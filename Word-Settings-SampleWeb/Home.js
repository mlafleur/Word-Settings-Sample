/// <reference path="/Scripts/FabricUI/MessageBanner.js" />


(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Add a click event handler for the highlight button.
            $('#run-button').click(runButton);
        });
    };



    function runButton() {
        Word.run(function (context) {

            return context.sync()
                .then(function () {

                    Office.context.document.settings.refreshAsync(function () {
                        var foo = Office.context.document.settings.get('hello');
                        if (!foo) {
                            Office.context.document.settings.set('hello', 'world');
                            Office.context.document.settings.saveAsync(function (asyncResult) {
                                $('#content').html('Settings saved with status: ' + asyncResult.status);
                            });
                        }
                        else {
                            $('#content').html('Value found: ' + foo);
                        }
                    });

                })
                .then(context.sync);
        });
    }
})();
