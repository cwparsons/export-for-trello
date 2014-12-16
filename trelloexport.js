/*!
 * TrelloExport
 * https://github.com/llad/export-for-trello
 *
 * Credit:
 * Started from: https://github.com/Q42/TrelloScrum
 */

/*jslint browser: true, devel: false*/

(function($, xlsx, saveAs) {
    // Variables
    var $excel_btn,
        addInterval,
        byteString,
        columnHeadings = [
            'Board ID',
            'Board Name',
            'List ID',
            'List Name',
            'Card #',
            'Card Name',
            'Card URL',
            'Card Description',
            'Labels',
            'Members',
            'Due Date',
            'Last Activity Date',
            'Last Activity',
            'Attachment Count',
            'Attachment Links',
            'Checklist Item Total Count',
            'Checklist Item Completed Count',
            'Vote Count',
            'Comment Count',
            'Archived'
        ];

    window.URL = window.webkitURL || window.URL;

    function createExcelExport() {
        // RegEx to find the points for users of TrelloScrum
        var pointReg = /[\(](\x3f|\d*\.?\d+)([\)])\s?/m;
        var jsonLink = $('.js-export-json[href]').attr('href');

        $.getJSON(jsonLink, function (data) {
            var file = {
                worksheets: [[], []], // Worksheets has one empty worksheet (array)
                creator: 'TrelloExport',
                created: new Date(),
                lastModifiedBy: 'TrelloExport',
                modified: new Date(),
                activeWorksheet: 0
            },

            // Setup the active list and cart worksheet
            worksheet = file.worksheets[0],
            wArchived = file.worksheets[1],
            buffer,
            i,
            ia,
            blob,
            board_title;

            // Set bold on column headers
            columnHeadings.forEach(function (element, index) {
                columnHeadings[index] = {
                    value: element,
                    bold: true
                };
            });

            worksheet.name = data.name.substring(0, 22);  // Over 22 chars causes Excel error, don't know why
            worksheet.data = [];
            worksheet.data.push([]);
            worksheet.data[0] = columnHeadings;

            // Setup the archive list and cart worksheet
            wArchived.name = 'Archived ' + data.name.substring(0, 22);
            wArchived.data = [];
            wArchived.data.push([]);
            wArchived.data[0] = columnHeadings;

            // This iterates through each list and builds the dataset
            $.each(data.lists, function (key, list) {
                // Tag archived lists
                if (list.closed) {
                    list.name = '[Archived] ' + list.name;
                }

                // Iterate through each card and transform data as needed
                $.each(data.cards, function (i, card) {
                    if (card.idList === list.id) {
                        var title = card.name,
                            parsed = title.match(pointReg),
                            points = parsed ? parsed[1] : '';

                        title = title.replace(pointReg, '');

                        // URL
                        var url = {
                            value: '<a href="' + card.shortUrl + '">' + card.shortUrl + '</a>'
                        };

                        var memberInitials = [];
                        $.each(card.idMembers, function (i, memberID) {
                            $.each(data.members, function (key, member) {
                                if (member.id === memberID) {
                                    memberInitials.push(member.initials);
                                }
                            });
                        });

                        // Get all labels
                        var labels = [];
                        $.each(card.labels, function (i, label) {
                            if (label.name) {
                                labels.push(label.name);
                            } else {
                                labels.push(label.color);
                            }
                        });

                        // Need to set dates to the Date type so xlsx.js sets the right datatype
                        var due = card.due || '';
                        if (due !== '') {
                            due = new Date(due);
                        }

                        // Attachments
                        var attachments = [];
                        if (data.attachments) {
                            data.attachments.forEach(function (element) {
                                attachments.push(element.url);
                            });
                        }

                        // Get activites by card ID
                        var activities = [];
                        if (data.actions) {
                            activities = data.actions.filter(function(element) {
                                return element.type === 'commentCard' && element.data.card.idShort === card.idShort;
                            });
                        }

                        var lastActivity = '';
                        if (activities.length && activities[0].data.text) {
                            lastActivity = '[' + activities[0].data.text + '] ' + activities[0].data.text;
                        }

                        var rowData = [
                            data.id,                        // 'Board ID',
                            data.name,                      // 'Board Name',
                            card.idList,                    // 'List ID',
                            list.name,                      // 'List Name',
                            card.idShort,                   // 'Card #',
                            title,                          // 'Card Name',
                            url,                            // 'Card URL',
                            card.desc,                      // 'Card Description',
                            labels.join(', '),              // 'Labels',
                            memberInitials.join(', '),      // 'Members',
                            due,                            // 'Due Date',
                            card.dateLastActivity,          // 'Last Activity Date',
                            lastActivity,                   // 'Last Activity',
                            card.badges.attachments,        // 'Attachment Count',
                            attachments.join(', '),         // 'Attachment Links',
                            card.badges.checkItems,         // 'Checklist Item Total Count',
                            card.badges.checkItemsChecked,  // 'Checklist Item Completed Count',
                            card.badges.votes,              // 'Vote Count',
                            card.badges.comments,           // 'Comment Count',
                            card.closed                     // 'Archived'
                        ];

                        // Writes all closed items to the Archived tab
                        // Note: Trello allows open cards on closed lists
                        if (list.closed || card.closed) {
                            var rArch = wArchived.data.push([]) - 1;
                            wArchived.data[rArch] = rowData;
                        } else {
                            var r = worksheet.data.push([]) - 1;
                            worksheet.data[r] = rowData;
                        }
                    }
                });
            });

            // We want just the base64 part of the output of xlsx.js
            // since we are not leveraging they standard transfer process.
            byteString = window.atob(xlsx(file).base64);
            buffer = new ArrayBuffer(byteString.length);
            ia = new Uint8Array(buffer);

            // Write the bytes of the string to an ArrayBuffer
            for (i = 0; i < byteString.length; i += 1) {
                ia[i] = byteString.charCodeAt(i);
            }

            // Create blob and save it using FileSaver.js
            blob = new Blob([ia], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            });

            board_title = data.name;
            saveAs(blob, board_title + '.xlsx');
            $('a.close-btn').eq(0).trigger('click');
        });

    }


    // Add a Export Excel button to the DOM and trigger export if clicked
    function addExportLink() {
        var $js_btn = $('.js-export-json[href]'); // Export JSON link

        // See if our Export Excel is already there
        if ($('.pop-over-list').find('.js-export-excel').length) {
            clearInterval(addInterval);
            return;
        }

        // The new link/button
        if ($js_btn.length) {
            $excel_btn = $('<a>')
                .attr({
                    'class': 'js-export-excel',
                    'href': '#',
                    'target': '_blank',
                    'title': 'Open downloaded file with Excel'
                })
                .text('Export Excel')
                .click(createExcelExport)
                .insertAfter($js_btn.parent())
                .wrap(document.createElement('li'));
        }
    }

    // On DOM load
    $(function () {
        // Look for clicks on the .js-share class, which is
        // the 'Share, Print, Export...' link on the board header option list
        $(document).on('mouseup', '.js-share', function () {
            addInterval = setInterval(addExportLink, 250);
        });
    });
})($, xlsx, saveAs);