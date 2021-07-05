const PS = new PerfectScrollbar("#cells", {
    wheelSpeed: 2,
    wheelPropagation: true,
});

function findRowCol(ele) {
    let idArray = $(ele).attr("id").split("-");
    let rowId = parseInt(idArray[1]);
    let colId = parseInt(idArray[3]);
    return [rowId, colId];
}

function calcColName(n) {
    let str = "";

    while (n > 0) {
        let rem = n % 26;
        if (rem == 0) {
            str = 'Z' + str;
            n = Math.floor((n / 26)) - 1;
        } else {
            str = String.fromCharCode((rem - 1) + 65) + str;
            n = Math.floor((n / 26));
        }
    }
    return str;
}

for (let i = 1; i <= 100; i++) {
    let str = calcColName(i);
    $("#columns").append(`<div class="column-name">${str}</div>`);
    $("#rows").append(`<div class="row-name">${i}</div>`);
}

//perfect scrollbar
$("#cells").scroll(function () {
    $("#columns").scrollLeft(this.scrollLeft);
    $("#rows").scrollTop(this.scrollTop);
})

let saved = true;
let cellData = { "Sheet1": {} };
let selectedSheet = "Sheet1";
let totalSheets = 1;
let lastelyAddedSheetNumber = 1;
let mousemoved = false;
let startCellStored = false;
let startCell;
let endCell;
let defaultProperties = {
    "font-family": "Noto Sans",
    "font-size": 14,
    "text": "",
    "bold": false,
    "italic": false,
    "underlined": false,
    "alignment": "left",
    "bgcolor": "#fff",
    "color": "#444",
    "border": "none",
    "formula": "",
    "upStream": [],
    "downStream": []
}

function loadNewSheet() {
    $("#cells").text("");
    for (let i = 1; i <= 100; i++) {
        let row = $(`<div class="cell-row"></div>`);
        for (let j = 1; j <= 100; j++){
            row.append(`<div id="row-${i}-col-${j}" class="input-cell" contenteditable="false"></div>`)
        }
        $("#cells").append(row);
    }
    addEventsToCells();
    addSheetTabEventListeners();
}

loadNewSheet();

//Add Events to cells
function addEventsToCells() {
    $(".input-cell").dblclick(function () {
        $(this).attr("contenteditable", "true");
        $(this).focus();
    })
    
    $(".input-cell").blur(function () { debugger
        $(this).attr("contenteditable", "false");
        let [rowId, colId] = findRowCol(this);
        updateCellData("text", $(this).text());
        console.log(cellData[selectedSheet][rowId - 1][colId - 1]);
        if (cellData[selectedSheet][rowId - 1][colId - 1].formula != "") {
            updateStreams(this, []);
        }
        cellData[selectedSheet][rowId - 1][colId - 1].formula = "";
        let selfColCode = calcColName(colId);
        evalFormula(selfColCode + rowId);
    })
    
    $(".input-cell").click(function (e) {
        let [rowId, colId] = findRowCol(this);
        let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRight(rowId, colId);
    
        if ($(this).hasClass("selected") && e.ctrlKey) {
            unselectCell(this, e, topCell, bottomCell, leftCell, rightCell);
        }
        else {
            selectCell(this, e, topCell, bottomCell, leftCell, rightCell);
            printSelectedCell(rowId, colId);
            selectRowCol(e, rowId - 1, colId - 1);
        }
    })
    
    $(".input-cell").mousemove(function (e) {
        e.preventDefault();
        if (e.buttons == 1 && !e.ctrlKey) {
            $(".input-cell.selected").removeClass("selected top-selected  bottom-selected left-selected right-selected");
            $(".row-name, .column-name").removeClass("selected");
            mousemoved = true;
            if (!startCellStored) {
                let [rowId, colId] = findRowCol(e.target);
                startCell = { rowId: rowId, colId: colId };
                startCellStored = true;
            }
            else {
                let [rowId, colId] = findRowCol(e.target);
                endCell = { rowId: rowId, colId: colId };
                selectAllBetweenRange(startCell, endCell);
                

                if (startCell.rowId >= endCell.rowId) {
                    if (startCell.colId >= endCell.colId) {
                        $("#selected-cell").text((startCell.rowId - endCell.rowId + 1) + "R X " + (startCell.colId - endCell.colId + 1) + "C");
                    }
                    else {
                        $("#selected-cell").text((startCell.rowId - endCell.rowId + 1) + "R X " + (endCell.colId - startCell.colId + 1) + "C");
                    }
                }
                else {
                    if (startCell.colId >= endCell.colId) {
                        $("#selected-cell").text((endCell.rowId - startCell.rowId + 1) + "R X " + (startCell.colId - endCell.colId + 1) + "C");
                    }
                    else {
                        $("#selected-cell").text((endCell.rowId - startCell.rowId + 1) + "R X " + (endCell.colId - startCell.colId + 1) + "C");
                    }
                }
            }
        }
        else if (e.buttons == 0 && mousemoved) {
            printSelectedCell(startCell.rowId, startCell.colId);
            startCellStored = false;
            mousemoved = false;
        }
    })
}
//

function getTopBottomLeftRight(rowId, colId) {
    let topCell = $(`#row-${rowId - 1}-col-${colId}`);
    let bottomCell = $(`#row-${rowId + 1}-col-${colId}`);
    let leftCell = $(`#row-${rowId}-col-${colId - 1}`);
    let rightCell = $(`#row-${rowId}-col-${colId + 1}`);

    return [topCell, bottomCell, leftCell, rightCell];
}

function unselectCell(ele, e, topCell, bottomCell, leftCell, rightCell) {
    if (e.ctrlKey && $(ele).attr("contenteditable") == "false") {
        
        if ($(ele).hasClass("top-selected")) {
            topCell.removeClass("bottom-selected");
        }
        
        if ($(ele).hasClass("bottom-selected")) {
            bottomCell.removeClass("top-selected");
        }
        
        if ($(ele).hasClass("left-selected")) {
            leftCell.removeClass("right-selected");
        }
        
        if ($(ele).hasClass("right-selected")) {
            rightCell.removeClass("left-selected");
        }
        $(ele).removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
}

function selectCell(ele, e, topCell, bottomCell, leftCell, rightCell, mouseSelection) {
    if (e.ctrlKey || mouseSelection) {
        
        // top selected or not
        let topSelected;
        if (topCell) {
            topSelected = topCell.hasClass("selected");
        }
        // bottom selected or not
        let bottomSelected;
        if (bottomCell) {
            bottomSelected = bottomCell.hasClass("selected");
        }

        // left selected or not
        let leftSelected;
        if (leftCell) {
            leftSelected = leftCell.hasClass("selected");
        }
        // right selected or not
        let rightSelected;
        if (rightCell) {
            rightSelected = rightCell.hasClass("selected");
        }

        if (topSelected) {
            topCell.addClass("bottom-selected");
            $(ele).addClass("top-selected");
        }

        if (leftSelected) {
            leftCell.addClass("right-selected");
            $(ele).addClass("left-selected");
        }

        if (rightSelected) {
            rightCell.addClass("left-selected");
            $(ele).addClass("right-selected");
        }

        if (bottomSelected) {
            bottomCell.addClass("top-selected");
            $(ele).addClass("bottom-selected");
        }
    }
    else {
        $(".input-cell").removeClass("selected top-selected bottom-selected left-selected right-selected");
    }
    $(ele).addClass("selected");
    changeHeader(findRowCol(ele));
}

function printSelectedCell(rowId, colId) {
    let str = calcColName(colId);

    $("#selected-cell").text(str + rowId);
}

function changeHeader([rowId, colId]) {
    let data;
    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        data = cellData[selectedSheet][rowId - 1][colId - 1];
    }
    else {
        data = defaultProperties;
    }
    $("#font-family").val(data["font-family"]);
    $("#font-family").css("font-family", data["font-family"]);
    $("#font-size").val(data["font-size"]);
    $(".alignment.selected").removeClass("selected");
    $(`.alignment[data-type=${data.alignment}]`).addClass("selected");
    addRemoveSelectFromFontStyle(data, "bold");
    addRemoveSelectFromFontStyle(data, "italic");
    addRemoveSelectFromFontStyle(data, "underlined");
    $("#fill-color-icon").css("border-bottom", `4px solid ${data.bgcolor}`);
    $("#text-color-icon").css("border-bottom", `4px solid ${data.color}`);
    $("#border").val(data["border"]);
    console.log(data.formula);
    $("#formula-input").text(data.formula);
}

function addRemoveSelectFromFontStyle(data, property) {
    if (data[property]) {
        $(`#${property}`).addClass("selected");
    }
    else {
        $(`#${property}`).removeClass("selected");
    }
}

function selectAllBetweenRange(start, end) {
    for (let i = (start.rowId < end.rowId ? start.rowId : end.rowId); i <= (start.rowId < end.rowId ? end.rowId : start.rowId); i++){
        for (let j = (start.colId < end.colId ? start.colId : end.colId); j <= (start.colId < end.colId ? end.colId : start.colId); j++){
            let [topCell, bottomCell, leftCell, rightCell] = getTopBottomLeftRight(i, j);
            selectCell($(`#row-${i}-col-${j}`), {}, topCell, bottomCell, leftCell, rightCell, true);
            selectRowCol({}, i - 1, j - 1, true);
        }
    }
}

function selectRowCol(e, rowIdx, colIdx, mouseSelection) {
    if (e.ctrlKey || mouseSelection) {
        $($(".row-name")[rowIdx]).addClass("selected");
        $($(".column-name")[colIdx]).addClass("selected");
    }
    else {
        $(".row-name, .column-name").removeClass("selected");
    }

    $($(".row-name")[rowIdx]).addClass("selected");
    $($(".column-name")[colIdx]).addClass("selected");
}

$(".menu-selector").change(function (e) {
    let value = $(this).val();
    let key = $(this).attr("id");
    if (key == 'font-family') {
        $("#font-family").css(key, value);
    }
    if (!isNaN(value)) {
        value = parseInt(value);
    }

    $(".input-cell.selected").css(key, value);
    updateCellData(key, value);
})

$("#border").change(function (e) {
    let value = $(this).val();
    $(".input-cell.selected").css("border", "");
    if (value == "none") {
        $(".input-cell.selected").removeClass("border-outer border-left border-right border-top border-bottom");
    }
    else if (value == "outer") {
        $(".input-cell.selected").addClass("border-outer");
    }
    else {
        $(".input-cell.selected").css(`border-${value}`, "3px solid #444");
    }
    updateCellData("border", value);
})

$("#search-box").click(function (e) {
    $(".search-options-modal").remove();
    $(this).addClass("selected");
    let searchBoxModal = $(`<div class="search-options-modal">
                            <div class="search-options-modal-title">Find & Select</div>
                            <div class="search-options">
                                <div class="search-option search">
                                    <div class="material-icons search-icon">search</div>
                                    <div>Find</div>
                                </div>
                                <div class="search-option replace">
                                    <div class="material-icons replace-icon">edit</div>
                                    <div>Replace</div>
                                </div>
                            </div>
                        </div>`)
    $(".container").append(searchBoxModal);
    let searchPosition = $(this).position();
    $(".search-options-modal").css({ "top": searchPosition.top + 37, "left": searchPosition.left });

    $(".search").click(function () {
        let searchModal = $(`<div class="sheet-find-modal">
                                <div class="sheet-modal-title">
                                    <span>Find</span>
                                    <span class="material-icons close-modal">close</span>
                                </div>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">Find what:</span>
                                    <input class="sheet-modal-input" type="text">
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button ok-button find">Find</div>
                                    <div class="button ok-button findAll">Find All</div>
                                </div>
                            </div>`)
        $(".container").append(searchModal);
        $(".sheet-modal-input").focus();
        $(".sheet-find-modal").draggable();

        $(".sheet-modal-input").click(function (e) {
            $(".error").remove();
        });
        
        $(".findAll").click(function (e) {
            searchData(e);
        });

        $(".find").click(function (e) {
            searchData(e);
        });

        $(".close-modal").click(function (e) {
            $(".sheet-find-modal").remove();
        });
    });

    $(".replace").click(function (e) {
        let replaceModal = $(`<div class="sheet-replace-modal">
                                <div class="sheet-modal-title">
                                    <span>Find and Replace</span>
                                    <span class="material-icons close-modal">close</span>
                                </div>
                                <div class="sheet-modal-input-container">
                                    <span class="sheet-modal-input-title">Find what:</span>
                                    <input class="sheet-modal-input" type="text">
                                    <span class="sheet-modal-input-title">Replace with:</span>
                                    <input class="sheet-modal-input" type="text">
                                </div>
                                <div class="sheet-modal-confirmation">
                                    <div class="button ok-button find">Find</div>
                                    <div class="button ok-button replace">Replace</div>
                                    <div class="button ok-button replaceAll">Replace All</div>
                                </div>
                            </div>`);
        $(".container").append(replaceModal);
        $($(".sheet-modal-input")[0]).focus();
        $(".sheet-replace-modal").draggable();

        $($(".sheet-modal-input")[0]).click(function (e) {
            $(".error").remove();
        });
        
        $(".replaceAll").click(function (e) {
            $(".error").remove();
            replaceData(e);
        });
        $(".replace").click(function (e) {
            $(".error").remove();
            replaceData(e);
        });
        $(".find").click(function (e) {
            replaceData(e);
        });

        $(".close-modal").click(function (e) {
            $(".sheet-replace-modal").remove();
        });
    })
})

$("#bold").click(function (e) {
    setFontStyle(this, "bold", "font-weight", "bold");
})

$("#italic").click(function (e) {
    setFontStyle(this, "italic","font-style", "italic");
})

$("#underlined").click(function (e) {
    setFontStyle(this, "underlined","text-decoration", "underline"); 
})

function setFontStyle(ele, property, key, value) {
    if ($(ele).hasClass("selected")) {
        $(ele).removeClass("selected");
        $(".input-cell.selected").css(key, "");
        updateCellData(property, false);
    }
    else {
        $(ele).addClass("selected");
        $(".input-cell.selected").css(key, value);
        updateCellData(property, true);
    }
}

$(".alignment").click(function (e) {
    $(".alignment.selected").removeClass("selected");
    $(this).addClass("selected");
    let alignment = $(this).attr("data-type");
    $(".input-cell.selected").css("text-align", alignment);
    updateCellData("alignment", alignment);
})

let searchValue;
let searchedData = [];
let searchIdx = 0;
let replaceIdx = 0;

function searchData(e) {
    let data = cellData[selectedSheet];
    if ($(".sheet-modal-input").val()) {
        if (searchValue != $(".sheet-modal-input").val()) {
            searchedData = [];
            searchValue = $(".sheet-modal-input").val();
            let rows = Object.keys(data);
            for (let i of rows) {
                let cols = Object.keys(data[i]);
                for (let j of cols) {
                    if (searchValue == data[i][j]["text"]) {
                        let searchedRow = parseInt(i) + 1;
                        let searchedCol = parseInt(j) + 1;
                        searchedData.push($(`#row-${searchedRow}-col-${searchedCol}`));
                    }
                }
            }
        }

        $(".input-cell.selected").removeClass("selected");
        if (e.currentTarget.textContent == "Find All") {
            for (let i of searchedData) {
                i.addClass("selected");
            }
        }
        else {
            if (searchIdx < searchedData.length) {
                searchedData[searchIdx].addClass("selected");
                searchIdx++;
            }
            else {
                searchedData[0].addClass("selected");
                searchIdx = 1;
            }
        }
    }
    else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class = "error"><span class="material-icons error-icon">error_outline</span> Search value does not exists </div>
        `);
    }
}

function replaceData(e) {
    let data = cellData[selectedSheet];
    let replaceValue = $($(".sheet-modal-input")[1]).val();
    if ($(".sheet-modal-input").val()) {
        if (searchValue != $(".sheet-modal-input").val()) {
            searchedData = [];
            searchValue = $(".sheet-modal-input").val();
            let rows = Object.keys(data);
            for (let i of rows) {
                let cols = Object.keys(data[i]);
                for (let j of cols) {
                    if (searchValue == data[i][j]["text"]) {
                        let searchedRow = parseInt(i) + 1;
                        let searchedCol = parseInt(j) + 1;
                        searchedData.push($(`#row-${searchedRow}-col-${searchedCol}`));
                    }
                }
            }
        }
        
        $(".input-cell.selected").removeClass("selected");
        if (e.currentTarget.textContent == "Replace All") {
            if (searchedData.length == 0) {
                $(".error").remove();
                $(".sheet-modal-input-container").append(`<div class = "error"><span class="material-icons error-icon">error_outline</span> Cannot find anything to remove </div>`);
            }
            else {
                for (let i of searchedData) {
                    i.addClass("selected");
                    i.text(replaceValue);
                    let [rowId, colId] = findRowCol(i);
                    cellData[selectedSheet][rowId - 1][colId - 1]["text"] = replaceValue;
                    searchedData = [];
                }
            }
        }
        else if (e.currentTarget.textContent == "Replace") {
            if (searchedData.length == 0) {
                $(".error").remove();
                $(".sheet-modal-input-container").append(`<div class = "error"><span class="material-icons error-icon">error_outline</span> Cannot find anything to remove </div>`);
            }
            else {
                if (replaceIdx < searchedData.length) {
                    searchedData[replaceIdx].addClass("selected");
                    searchedData[replaceIdx].text(replaceValue);
                    let [rowId, colId] = findRowCol(searchedData[replaceIdx]);
                    cellData[selectedSheet][rowId - 1][colId - 1]["text"] = replaceValue;
                    searchedData.splice(replaceIdx, 1);
                    searchIdx = replaceIdx;
                }
            }
        }
        else {
            if (searchIdx < searchedData.length) {
                searchedData[searchIdx].addClass("selected");
                replaceIdx = searchIdx;
                searchIdx++;
            }
            else {
                searchedData[0].addClass("selected");
                replaceIdx = 0;
                searchIdx = 1;
            }
        }
    }
    else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class = "error"><span class="material-icons error-icon">error_outline</span> Search value does not exists </div>
        `);
    }
}

function updateCellData(property, value) {
    let prevCellData = JSON.stringify(cellData);
    if (value != defaultProperties[property]) {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            if (cellData[selectedSheet][rowId - 1] == undefined) {
                cellData[selectedSheet][rowId - 1] = {};
                cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties }  //{...array/object} is sparse array or object used to make a copy of array or object
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
            }
            else {
                if (cellData[selectedSheet][rowId - 1][colId - 1] == undefined) {
                    cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties };
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
                else {
                    cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                }
            }
        });
    }
    else {
        $(".input-cell.selected").each(function (index, data) {
            let [rowId, colId] = findRowCol(data);
            if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
                cellData[selectedSheet][rowId - 1][colId - 1][property] = value;
                if (JSON.stringify(cellData[selectedSheet][rowId - 1][colId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][rowId - 1][colId - 1];
                    if(Object.keys(cellData[selectedSheet][rowId - 1]).length == 0) {
                        delete cellData[selectedSheet][rowId - 1];
                    }
                }
            }
        });
    }

    if (saved && JSON.stringify(cellData) != prevCellData) {
        saved = false;
    }
}

$(".color-pick").colorPick({
    'initialColor': '#TYPECOLOR',
    'allowRecent': true,
    'recentMax': 5,
    'allowCustomColor': true,
    'palette': ["#1abc9c", "#16a085", "#2ecc71", "#27ae60", "#3498db", "#2980b9", "#9b59b6", "#8e44ad", "#34495e", "#2c3e50", "#f1c40f", "#f39c12", "#e67e22", "#d35400", "#e74c3c", "#c0392b", "#ecf0f1", "#bdc3c7", "#95a5a6", "#7f8c8d"],
    'onColorSelected': function() {
        if (this.color != "#TYPECOLOR") {
            if (this.element.attr("id") == "fill-color") {
                $("#fill-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("background-color", this.color);
                updateCellData("bgcolor", this.color);
            }
            else {
                $("#text-color-icon").css("border-bottom", `4px solid ${this.color}`);
                $(".input-cell.selected").css("color", this.color);
                updateCellData("color", this.color);
            }
        }
    }
});

$("#fill-color-icon, #text-color-icon").click(function (e) {
    setTimeout(() => {
        $(this).parent().click();
    }, 5);
});

$(".container").click(function(e) {
    $(".sheet-options-modal").remove();

    if ($(".sheet-list-modal").hasClass("active")) {
        $(".sheet-list-modal").remove();
    }
    else {
        $(".sheet-list-modal").addClass("active");
    }

    if ($(".search-options-modal").hasClass("active")) {
        $(".search-options-modal").removeClass("active");
        $("#search-box").removeClass("selected");
        $(".search-options-modal").remove();
    }
    else {
        $(".search-options-modal").addClass("active");
    }
});

function selectSheet(ele) {
    $(".sheet-tab.selected").removeClass("selected");
    $(ele).addClass("selected");
    emptySheet();
    selectedSheet = $(ele).text();
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
    loadSheet();
}

function emptySheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text("");
            cell.css({
                "font-family" : "Noto Sans",
                "font-size" : 14,
                "background-color" : "#fff",
                "color": "#444",
                "font-weight" : "",
                "font-style" : "",
                "text-decoration" : "",
                "text-align": "left",
                "border": ""
            });
            cell.removeClass("border-outer border-top border-left border-bottom border-right");
        }
    }
}

function loadSheet() {
    let data = cellData[selectedSheet];
    let rowKeys = Object.keys(data);
    for(let i of rowKeys) {
        let rowId = parseInt(i);
        let colKeys = Object.keys(data[rowId]);
        for (let j of colKeys) {
            let colId = parseInt(j);
            let cell = $(`#row-${rowId + 1}-col-${colId + 1}`);
            cell.text(data[rowId][colId].text);
            cell.css({
                "font-family" : data[rowId][colId]["font-family"],
                "font-size" : data[rowId][colId]["font-size"] + "px",
                "background-color" : data[rowId][colId]["bgcolor"],
                "color": data[rowId][colId].color,
                "font-weight" : data[rowId][colId].bold ? "bold" : "",
                "font-style" : data[rowId][colId].italic ? "italic" : "",
                "text-decoration" : data[rowId][colId].underlined ? "underline" : "",
                "text-align" : data[rowId][colId].alignment 
            });
            data[rowId][colId]["border"] == "none" ? cell.css("border", "") : cell.addClass(`border-${data[rowId][colId]["border"]}`);
        }
    }
}

$(".add-sheet").click(function (e) {
    emptySheet();
    totalSheets++;
    lastelyAddedSheetNumber++;
    while (Object.keys(cellData).includes("Sheet" + lastelyAddedSheetNumber)) {
        lastelyAddedSheetNumber++;
    }
    cellData[`Sheet${lastelyAddedSheetNumber}`] = {};
    selectedSheet = `Sheet${lastelyAddedSheetNumber}`;
    $(".sheet-tab.selected").removeClass("selected");
    $(".sheet-tab-container").append(
        `<div class="sheet-tab selected">Sheet${lastelyAddedSheetNumber}</div>`
    );
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
    addSheetTabEventListeners();
    $("#row-1-col-1").click();
    saved = false;
});

//add sheet tab event listeners
function addSheetTabEventListeners() {
    $(".sheet-tab.selected").bind("contextmenu", function (e) {
        e.preventDefault();
        $(".sheet-options-modal").remove();
        $(".sheet-list-modal").remove();
        let modal = $(
            `<div class="sheet-options-modal">
                <div class="option sheet-rename">Rename</div>
                <div class="option sheet-delete" >Delete</div>
            </div>`
        )
        $(".container").append(modal);
        $(".sheet-options-modal").css({ "bottom": 0.04 * $(".container").height(), "left": e.pageX });
        if (totalSheets == 1) {
            $(".sheet-delete").addClass("disabled");
        }
        $(".sheet-rename").click(function (e) {
            let renameModal = `<div class="sheet-modal-parent">
                                    <div class="sheet-rename-modal">
                                        <div class="sheet-modal-title">
                                            <span>Rename Sheet</span>
                                        </div>
                                        <div class="sheet-modal-input-container">
                                            <span class="sheet-modal-input-title">Rename sheet to:</span>
                                            <input class="sheet-modal-input" type="text">
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button ok-button">OK</div>
                                            <div class="button cancel-button">Cancel</div>
                                        </div>
                                    </div>
                                </div>`;
            
            $(".container").append(renameModal);
            $(".sheet-modal-input").focus();
    
            $(".cancel-button").click(function () {
                $(".sheet-modal-parent").remove();
            })
    
            $(".ok-button").click(function () {
                renameSheet();
            })
    
            $(".sheet-modal-input").keypress(function (e) {
                if (e.key == "Enter") {
                    renameSheet();
                }
            })
        });
    
        if (!$(".sheet-delete").hasClass("disabled")) {
            $(".sheet-delete").click(function (e) {
                let deleteModal = `<div class="sheet-modal-parent">
                                        <div class="sheet-delete-modal">
                                            <div class="sheet-modal-title">
                                                <span>Sheet Name</span>
                                            </div>
                                            <div class="sheet-modal-detail-container">
                                                <span class="sheet-modal-detail-title">Are you Sure?</span>
                                            </div>
                                            <div class="sheet-modal-confirmation">
                                                <div class="button delete-button">
                                                    <span class="material-icons delete-icon">delete</span>
                                                    Delete
                                                </div>
                                                <div class="button cancel-button">Cancel</div>
                                            </div>
                                        </div>
                                    </div>`;
                
                $(".container").append(deleteModal);
        
                $(".cancel-button").click(function () {
                    $(".sheet-modal-parent").remove();
                })
        
                $(".delete-button").click(function (e) {
                    deleteSheet(e);
                })
            });
        }
    
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
        }
    });

    $(".sheet-tab.selected").click(function (e) {
        if (!$(this).hasClass("selected")) {
            selectSheet(this);
            $("#row-1-col-1").click();
        }
    });
}

function renameSheet() {
    let newSheetName = $(".sheet-modal-input").val();
    if (newSheetName && !Object.keys(cellData).includes(newSheetName)) {
        let newCellData = {};
        for (let i of Object.keys(cellData)) {
            if (i == selectedSheet) {
                newCellData[newSheetName] = cellData[i];
            }
            else {
                newCellData[i] = cellData[i];
            }
        }
        cellData = newCellData;
        selectedSheet = newSheetName;
                    
        $(".sheet-tab.selected").text(newSheetName);
        $(".sheet-modal-parent").remove();
        saved = false;
    }
    else {
        $(".error").remove();
        $(".sheet-modal-input-container").append(`
            <div class = "error"><span class="material-icons error-icon">error_outline</span> Sheet Name is not valid or Sheet already exists </div>
        `);
    }
}

function deleteSheet(ele) {
    $(".sheet-modal-parent").remove();
    let keyArray = Object.keys(cellData);
    let selectedSheetIndex = keyArray.indexOf(selectedSheet);
    let currentSelectedSheet = $(".sheet-tab.selected");
    if (selectedSheetIndex == 0) {
        selectSheet(currentSelectedSheet.next()[0]);
    }
    else {
        selectSheet(currentSelectedSheet.prev()[0]);
    }
    delete cellData[currentSelectedSheet.text()];
    currentSelectedSheet.remove();
    totalSheets--;
    saved = false;
}

$(".left-scroller").click(function (e) {

    let keysArray = Object.keys(cellData);  
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != 0) {
        selectSheet($(".sheet-tab.selected").prev()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
})

$(".right-scroller").click(function (e) {
    let keysArray = Object.keys(cellData);
    let selectedSheetIndex = keysArray.indexOf(selectedSheet);
    if (selectedSheetIndex != (keysArray.length - 1)) {
        selectSheet($(".sheet-tab.selected").next()[0]);
    }
    $(".sheet-tab.selected")[0].scrollIntoView({ block: "nearest" });
})

$(".sheet-menu").click(function (e) {
    $(".sheet-list-modal").remove();
    let sheetMenu = $(`<div class="sheet-list-modal">
                        </div>`);
    
    $(".container").append(sheetMenu);
    $(".sheet-list-modal").css({ "bottom": 0.04 * $(".container").height(), "left": e.pageX });
    let sheetArray = $(".sheet-tab");
    for (let i = 0; i < sheetArray.length; i++) {
        sheetMenu.append(`<div class="sheet-list" id=${i + 1}>${sheetArray[i].textContent}</div>`)
    }

    $(".sheet-list").click(function (e) {
        let sheetIndex = $(this).attr("id") - 1;
        selectSheet(sheetArray[sheetIndex]);
    })
    
});

let sheetZoom = 1;
let sheetZoomPercentage = 100;
$("#zoom-percentage").text(sheetZoomPercentage + "%");

$("#zoom-in").click(function (e) {
    if (sheetZoomPercentage <= 200) {
        sheetZoom += 0.1;
        sheetZoomPercentage += 10;
        $(".data-container").css("zoom", sheetZoom);
        $("#zoom-percentage").text(sheetZoomPercentage + "%");
    }
})

$("#zoom-out").click(function (e) {
    if (sheetZoomPercentage >= 20) {
        sheetZoom -= 0.1;
        sheetZoomPercentage -= 10;
        $(".data-container").css("zoom", sheetZoom);
        $("#zoom-percentage").text(sheetZoomPercentage + "%");
    }
})

$("#menu-file").click(function () {
    let fileModal = $(`<div class="file-modal">
                        <div class="file-options-modal">
                            <div class="close">
                                <div class="material-icons close-icon">arrow_circle_down</div>
                                <div>Close</div>
                            </div>
                            <div class="new">
                                <div class="material-icons new-icon">insert_drive_file</div>
                                <div>New</div>
                            </div>
                            <div class="open">
                                <div class="material-icons open-icon">folder_open</div>
                                <div>Open</div>
                            </div>
                            <div class="save">
                                <div class="material-icons save-icon">save</div>
                                <div>Save</div>
                            </div>
                        </div>
                        <div class="file-recent-modal">
                            <div class="file-recent-modal-title">Recent</div>
                            <div class="recent-files-container">
                                <div class="recent-file">
                                    <div class="recent-file-icon-container">
                                        <div class="material-icons recent-file-icon">description</div>
                                    </div>
                                    <div class="recent-file-title">File name</div>
                                </div>
                                <div class="recent-file">
                                    <div class="recent-file-icon-container">
                                        <div class="material-icons recent-file-icon">description</div>
                                    </div>
                                    <div class="recent-file-title">File name</div>
                                </div>
                                <div class="recent-file">
                                    <div class="recent-file-icon-container">
                                        <div class="material-icons recent-file-icon">description</div>
                                    </div>
                                    <div class="recent-file-title">File name</div>
                                </div>
                            </div>
                        </div>
                        <div class="file-transparent-modal"></div>
                    </div>`);
    $(".container").append(fileModal);
    fileModal.animate({
        "width": "100vw"
    }, 200);

    $(".close, .file-transparent-modal, .new, .save, .open").click(function (e) {
        fileModal.animate({
            "width": "0vw"
        }, 200);
        setTimeout(() => {
            fileModal.remove();
        }, 200);
    });

    $(".new").click(function (e) {
        if (saved) {
            newFile();
        }
        else {
            $(".container").append(`<div class="sheet-modal-parent">
                                    <div class="sheet-delete-modal">
                                        <div class="sheet-modal-title">
                                            <span>${$(".title-bar").text()}</span>
                                        </div>
                                        <div class="sheet-modal-detail-container">
                                            <span class="sheet-modal-detail-title">Do you want to save changes?</span>
                                        </div>
                                        <div class="sheet-modal-confirmation">
                                            <div class="button ok-button">Save</div>
                                            <div class="button cancel-button">No</div>
                                        </div>
                                    </div>
                                </div>`);
            
            $(".ok-button").click(function (e) {
                $(".sheet-modal-parent").remove();
                saveFile(true);
            })

            $(".cancel-button").click(function (e) {
                newFile();
            });
        }
    });

    $(".save").click(function (e) {
        saveFile();
    })

    $(".open").click(function (e) {
        openFile();
    });
});

function newFile() {
    emptySheet();
    $(".sheet-tab").remove();
    $(".sheet-tab-container").append(`<div class="sheet-tab selected">Sheet1</div>`);
    cellData = { "Sheet1": {} };
    selectedSheet = "Sheet1";
    totalSheets = 1;
    lastelyAddedSheetNumber = 1;
    addSheetTabEventListeners();
    $("#row-1-col-1").click();
}

function saveFile(createNewFile) {
    if (!saved) {
        $(".container").append(`<div class="sheet-modal-parent">
                                <div class="sheet-rename-modal">
                                    <div class="sheet-modal-title">
                                        <span>Save File</span>
                                    </div>
                                    <div class="sheet-modal-input-container">
                                        <span class="sheet-modal-input-title">File Name:</span>
                                        <input class="sheet-modal-input" type="text" value="${$(".title-bar").text()}" />
                                    </div>
                                    <div class="sheet-modal-confirmation">
                                        <div class="button ok-button">Save</div>
                                        <div class="button cancel-button">Cancel</div>
                                    </div>
                                </div>
                            </div>`);
        
        $(".ok-button").click(function (e) {
            fileDownload(createNewFile);
        });

        $(".sheet-modal-input").keypress(function (e) {
            if (e.key == "Enter") {
                fileDownload(createNewFile);
            }
        });

        $(".cancel-button").click(function (e) {
            $(".sheet-modal-parent").remove();
            if (createNewFile) {
                newFile();
            }
        });
    }
}

function fileDownload(createNewFile){
    let fileName = $(".sheet-modal-input").val();
    if (fileName) {
        let href = `data:application/json,${encodeURIComponent(JSON.stringify(cellData))}`;
        let a = $(`<a href=${href} download="${fileName}.json"></a>`);
        $(".container").append(a)
        a[0].click();
        a.remove();
        $(".sheet-modal-parent").remove();
        saved = true;
        if (createNewFile) {
            newFile();
        }
    }
}

function openFile() {
    let inputFile = $(`<input type="file" accept="application/json" />`);
    $(".container").append(inputFile);
    inputFile.click();
    inputFile.change(function (e) {
        let file = e.target.files[0];
        $(".title-bar").text(file.name.split(".json")[0]);
        let reader = new FileReader();
        reader.readAsText(file);
        reader.onload = function () {
            emptySheet();
            cellData = JSON.parse(reader.result);
            $(".sheet-tab").remove();
            let sheets = Object.keys(cellData);
            for (i of sheets) {
                $(".sheet-tab-container").append(`<div class="sheet-tab selected">${i}</div>`);
            }
            addSheetTabEventListeners();
            $(".sheet-tab").removeClass("selected");
            $($(".sheet-tab")[0]).addClass("selected");
            selectedSheet = sheets[0];
            totalSheets = sheets.length;
            lastelyAddedSheetNumber = sheets.length;
            loadSheet();
            inputFile.remove();
        }
    })
}

let clipBoard = { startCell: [], cellData: {} };
let contentCutted = false;
// let copiedCell;

$("#copy, #cut").click(function (e) {
    if ($(this).text() == 'content_cut') {
        contentCutted = true;
    }
    clipBoard.startCell = findRowCol($(".input-cell.selected")[0]);
    $(".input-cell.selected").each((index, data) => {
        let [rowId, colId] = findRowCol(data);
        if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
            if (!clipBoard.cellData[rowId]) {
                clipBoard.cellData[rowId] = {};
            }
            clipBoard.cellData[rowId][colId] = { ...cellData[selectedSheet][rowId - 1][colId - 1] };
        }
    })
})

$("#paste").click(function (e) {
    // copiedCell.removeClass("rotating-dashed");
    if (contentCutted) {
        emptySheet();
    }
    let startCell = findRowCol($(".input-cell.selected")[0]);
    let rows = Object.keys(clipBoard.cellData);
    for (let i of rows) {
        let cols = Object.keys(clipBoard.cellData[i]);
        for (let j of cols) {
            if (contentCutted) {
                delete cellData[selectedSheet][i - 1][j - 1];
                if (Object.keys(cellData[selectedSheet][i - 1]).length == 0) {
                    delete cellData[selectedSheet][i - 1];
                }
            }
            let rowDistance = parseInt(i) - parseInt(clipBoard.startCell[0]);
            let colDistance = parseInt(j) - parseInt(clipBoard.startCell[1]);
            if (!cellData[selectedSheet][startCell[0] + rowDistance - 1]) {
                cellData[selectedSheet][startCell[0] + rowDistance - 1] = {};
            }
            cellData[selectedSheet][startCell[0] + rowDistance - 1][startCell[1] + colDistance - 1] = { ...clipBoard.cellData[i][j] };
        }
    }
    loadSheet();
    if (contentCutted) {
        contentCutted = false;
        clipBoard = { startCell: [], cellData: {} };
    }
});

$("#formula-input").blur(function (e) {
    if ($(".input-cell.selected").length > 0) {
        let formula = $(this).text();
        $(".input-cell.selected").each(function (index, data) {
            let tempElements = formula.split(" ");
            let elements = [];
            for (let i of tempElements) {
                if (i.length > 1) {
                    i = i.replace("(", "");
                    i = i.replace(")", "");
                    elements.push(i);
                }
            }

            if (updateStreams(data, elements, false)) {
                let [rowId, colId] = findRowCol(data);
                cellData[selectedSheet][rowId - 1][colId - 1].formula = formula;
                let selfColCode = calcColName(colId);
                evalFormula(selfColCode + rowId);
            }
            else {
                alert("Formula is invalid!");
            }
        })
    }
    else {
        alert("Please select a cell first ti apply formula");
    }
})

function updateStreams(ele, elements, update, oldUpStream) {
    let [rowId, colId] = findRowCol(ele);
    let selfColCode = calcColName(colId);
    for(let i = 0; i < elements.length; i++) {
        if(checkForSelf(rowId,colId,elements[i])){
            return false;
        }
    }

    if (cellData[selectedSheet][rowId - 1] && cellData[selectedSheet][rowId - 1][colId - 1]) {
        let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        for (let i of downStream) {
            if (elements.includes(i)) {
                return false;
            }
        }

        for (let i of downStream) {
            let [calRowId, calColId] = calcSelfValue(i);
            updateStreams($(`#row-${calRowId}-col-${calColId}`)[0], elements, true, upStream);
        }
    }

    if (!cellData[selectedSheet][rowId - 1]) {
        cellData[selectedSheet][rowId - 1] = {};
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else if (!cellData[selectedSheet][rowId - 1][colId - 1]) {
        cellData[selectedSheet][rowId - 1][colId - 1] = { ...defaultProperties, "upStream": [...elements], "downStream": [] };
    } else {
        let upStream = [...cellData[selectedSheet][rowId - 1][colId - 1].upStream];
        if (update) {
            for (let i of oldUpStream) {
                let [calRowId, calColId] = calcSelfValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1][calColId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
                index = cellData[selectedSheet][rowId - 1][colId - 1].upStream.indexOf(i);
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.splice(index, 1);
            }

            for (let i of elements) {
                cellData[selectedSheet][rowId - 1][colId - 1].upStream.push(i);
            }
        } else {
            for (let i of upStream) {
                let [calRowId, calColId] = calcSelfValue(i);
                let index = cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.indexOf(selfColCode + rowId);
                cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.splice(index, 1);
                if (JSON.stringify(cellData[selectedSheet][calRowId - 1][calColId - 1]) == JSON.stringify(defaultProperties)) {
                    delete cellData[selectedSheet][calRowId - 1][calColId - 1];
                    if (Object.keys(cellData[selectedSheet][calRowId - 1][calColId - 1]).length == 0) {
                        delete cellData[selectedSheet][calRowId - 1];
                    }
                }
            }
            cellData[selectedSheet][rowId - 1][colId - 1].upStream = [...elements];
        }
    }

    for (let i of elements) {
        let [calRowId, calColId] = calcSelfValue(i);
        if (!cellData[selectedSheet][calRowId - 1]) {
            cellData[selectedSheet][calRowId - 1] = {};
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else if (!cellData[selectedSheet][calRowId - 1][calColId - 1]) {
            cellData[selectedSheet][calRowId - 1][calColId - 1] = { ...defaultProperties, "upStream": [], "downStream": [selfColCode + rowId] };
        } else {
            cellData[selectedSheet][calRowId - 1][calColId - 1].downStream.push(selfColCode + rowId);
        }
    }

    return true;
}

function calcSelfValue(ele) {
    console.log(ele);
    let calRowId, calColId;

    for (let i = 0; i < ele.length; i++){
        if (!isNaN(ele.charAt(i))) {
            let leftString = ele.substring(0, i);
            let rightString = ele.substring(i);
            calColId = calcColId(leftString);
            calRowId = parseInt(rightString);
            break;
        }
    }
    return [calRowId, calColId];
}

function checkForSelf(rowId, colId, ele) {
    let [calRowId, calColId] = calcSelfValue(ele);
    if (calRowId == rowId && calColId == colId) {
        return true;
    }
    
    return false;
}

function calcColId(str) {
    let place = str.length - 1;
    let total = 0;
    for (let i = 0; i < str.length; i++) {
        let charValue = str.charCodeAt(i) - 64;
        total += Math.pow(26, place) * charValue;
        place--;
    }
    return total;
}

function evalFormula(cell) { debugger
    let [rowId, colId] = calcSelfValue(cell);
    let formula = cellData[selectedSheet][rowId - 1][colId - 1].formula;
    // console.log(formula);
    if (formula != ""){
        let upStream = cellData[selectedSheet][rowId - 1][colId - 1].upStream;
        let upStreamValue = [];
        for (let i in upStream) {
            let [calRowId, calColId] = calcSelfValue(upStream[i]);
            let value;
            if (cellData[selectedSheet][calRowId - 1][calColId - 1].text == "") {
                value = "0";
            }
            else {
                value = cellData[selectedSheet][calRowId - 1][calColId - 1].text;
            }
            upStreamValue.push(value);
            console.log(upStreamValue);
            formula = formula.replace(upStream[i], upStreamValue[i]);
        }


        cellData[selectedSheet][rowId - 1][colId - 1].text = eval(formula);
        loadSheet();
    }

    let downStream = cellData[selectedSheet][rowId - 1][colId - 1].downStream;
    for (let i = downStream.length - 1; i >= 0; i--){
        evalFormula(downStream[i]);
    }
}