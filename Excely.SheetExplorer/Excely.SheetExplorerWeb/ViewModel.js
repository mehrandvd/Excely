function ViewModel() {
    var self = this;

    self.title = ko.observable("Sheets Tree View");
    self.sheets = ko.observableArray([]);
    self.searchText = ko.observable("");
    self.isBusy = ko.observable(false);
    self.filteredSheets = ko.pureComputed(function() {
        return self.sheets().filter(function(item) {
            if (item.sheetInfo.name && self.searchText())
                return item.sheetInfo.name.toLowerCase().indexOf(self.searchText().toLowerCase()) !== -1;
            return true;
        }).map(function(sheet) {
            return {
                sheet: sheet,
                isExpanded: ko.observable(false)
            }
        });
    });

    self.refreshSheets = function () {
        self.isBusy(true);
        Excel.run(function (ctx) {
            // Queue a command to write the sample data to the worksheet
            ctx.workbook.worksheets.load('items');
            var p = ctx.sync();
            p.then(function() {
                    var list = ctx.workbook.worksheets.items.map(function(item) {
                        return {
                            sheetInfo: item,
                            dependecies: ko.observableArray([]),
                            description: '',
                            loadedFormulas: ko.observableArray([]),
                        }
                    });

                    self.sheets(list);

                    function loadFormulas(item) {
                        Excel.run(function (ctxFormula) {

                            var sheetName = item.sheetInfo.name;
                            //var rangeAddress = "A1:GG60";
                            var worksheet = ctxFormula.workbook.worksheets.getItem(sheetName);
                            item.loadedRange = worksheet.getUsedRange();
                            item.loadedRange.load('formulas');

                            ctxFormula.sync().then(function () {
                                var formulas = [];

                                for (var i = 0; i < item.loadedRange.formulas.length ; i++) {
                                    var row = item.loadedRange.formulas[i];
                                    for (var j = 0; j < row.length ; j++) {
                                        var cell = row[j];
                                        if (cell && cell.length > 0 &&  cell[0] ==='=')
                                            formulas.push(cell);
                                    }
                                }
                                item.loadedFormulas(formulas);

                                var dependentSheets = [];

                                for (var i = 0; i < self.sheets().length; i++) {
                                    var sheet = self.sheets()[i];

                                    var someFormulaUsesSheet = formulas.some(function (item) {
                                        return item.indexOf(sheet.sheetInfo.name) > -1;
                                    });

                                    if (someFormulaUsesSheet) {
                                        dependentSheets.push({
                                            sheet: sheet,
                                            isExpanded: ko.observable(false)
                                        });

                                    }

                                    item.dependecies(dependentSheets);

                                }
                            });
                        });
                    };

                    for (var i = 0; i < list.length; i++) {
                        var item = list[i];
                        loadFormulas(item);
                    }
                    self.isBusy(false);
                    //return ctx.sync().then(function() {
                    //    //for (var i = 0; i < list.length; i++) {
                    //    //    var item = list[i];

                    //    //    item.loadedFormulas = item.loadedRange.forumlas;
                    //    //}

                    //    self.sheets(list);
                    //});

                    // Run the queued-up commands, and return a promise to indicate task completion
                    //return ctx.sync();
                })
                .catch(errorHandler);
        });
    }

    function populateDescription(sheetInfo) {
        console.log(sheetInfo);
    }

    self.activateWorksheet = function(selectedSheet) {
        Excel.run(function (ctx) {
            var wSheetName = selectedSheet.name;
            var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
            worksheet.activate();
            return ctx.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }


    function init() {
        self.refreshSheets();
    }

    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    init();

    return self;
}