
(function () {
    "use strict";

    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If not using Word 2016, use fallback logic.
            if (!Office.context.requirements.isSetSupported('WordApi', '1.1')) {
                $("#template-description").text("This sample displays the selected text.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selected text");
                
                $('#word-count-button').click(displaySelectedText);
                return;
            }

            $("#template-description").text("This app counts the frequency of each word in the text you have selected in the document.");
            $('#button-text').text("Word Frequency!");
            $('#button-desc').text("Gets the frequency of all words.");
            
            loadSampleData();

            // Add a click event handler for the word count button.
            $('#word-count-button').click(getWordFrequency);
        });
    };

    function loadSampleData() {
        // Run a batch operation against the Word object model.
        Word.run(function (context) {
            // Create a proxy object for the document body.
            var body = context.document.body;

            // Queue a commmand to clear the contents of the body.
            body.clear();
            // Queue a command to insert text into the end of the Word document body.
            body.insertText(
                "This is a sample text inserted in the document",
                Word.InsertLocation.end);

            // Synchronize the document state by executing the queued commands, and return a promise to indicate task completion.
            return context.sync();
        })
        .catch(errorHandler);
    }

    function loadResults(results) {
        Word.run(function (context) {
            const wordFreqtable = document.createElement("table");

            const sortedResults = Object.fromEntries(
                Object.entries(results).sort(([, a], [, b]) => b - a)
            );


            for (const [key, value] of Object.entries(sortedResults)) {
                if (value > 2) {
                    const row = document.createElement("tr");
                    const header = document.createElement("th");
                    const cell = document.createElement("td");

                    header.textContent = key;
                    cell.textContent = value;

                    row.appendChild(header);
                    row.appendChild(cell);
                    wordFreqtable.appendChild(row);
                }
            }
            document.body.appendChild(wordFreqtable)
        })
    }

    function getWordFrequency() {
        Word.run(function (context) {
            // Queue a command to get the current selection and then
            // create a proxy range object with the results.
            var range = context.document.body;


            // Queue a command to load the range selection result.
            context.load(range, 'text');
            
            // This variable will keep the frequency of all the words in the selection.
            var wordFreq = {};


            // Synchronize the document state by executing the queued commands
            // and return a promise to indicate task completion.
            return context.sync()
                .then(function () {
                    // Get the longest word from the selection.
                    var words = range.text.split(/\s+/);
                    // get this from a library?
                    const conjunctions = ["an", "by", "it", "we", "me", "was", "my", "at", "in", "is", "i", "with", "to", "be", "a", "the", "of", "and", "but", "or", "nor", "for", "yet", "so", "although", "because", "since", "unless", "while", "whereas", "if", "unless", "until", "after", "before", "when", "while", "once"];
                    for (let word of words) {
                        if (!conjunctions.includes(word.toLowerCase())) {
                            wordFreq[word] = (wordFreq[word] || 0) + 1;
                        }
                    }
                })
                .then(context.sync)
                .then(function () {
                    // Queue a command to print the word frequency results.
                    loadResults(wordFreq);
                })
                .then(context.sync);
        })
        .catch(errorHandler);
    } 


    function displaySelectedText() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error:', result.error.message);
                }
            });
    }

    //$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
    function errorHandler(error) {
        // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
        showNotification("Error:", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
