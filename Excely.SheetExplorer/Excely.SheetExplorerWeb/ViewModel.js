function ViewModel() {
    var self = this;

    self.title = ko.observable("Sheets Tree View");
    self.sheets = ko.observableArray([]);
    self.searchText = ko.observable("");

    self.filteredSheets = ko.pureComputed(function() {
        return self.sheets().filter(function(item) {
            if (item.name && self.searchText())
                return item.name.toLowerCase().indexOf(self.searchText().toLowerCase()) !== -1;
            return true;
        });
    });

    self.refreshSheets = function () {
        Excel.run(function (ctx) {
                // Queue a command to write the sample data to the worksheet
                ctx.workbook.worksheets.load('items');
                var p = ctx.sync();
                p.then(function () {
                    var list = ctx.workbook.worksheets.items;
                    self.sheets(list);
                });

                // Run the queued-up commands, and return a promise to indicate task completion
                return ctx.sync();
            })
            .catch(errorHandler);
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