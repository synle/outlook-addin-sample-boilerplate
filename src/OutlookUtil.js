const OutlookUtil = {};

let _resolvePromiseDoneInit;
const promiseDone = new Promise((resolve) => {
    _resolvePromiseDoneInit = resolve;
});


// office item
let item;

OutlookUtil.initialize = () => promiseDone;


// Get the email addresses of all the recipients of the composed item.
OutlookUtil.getAllRecipients = () => new Promise((resolve) => {
  // Local objects to point to recipients of either
  // the appointment or message that is being composed.
  // bccRecipients applies to only messages, not appointments.
    let toRecipients,
        ccRecipients,
        bccRecipients;
  // Verify if the composed item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    } else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

  // Use asynchronous method getAsync to get each type of recipients
  // of the composed item. Each time, this example passes an anonymous
  // callback function that doesn't take any parameters.
    toRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    // write(asyncResult.error.message);
        } else {
    // Async call to get to-recipients of the item completed.
    // Display the email addresses of the to-recipients.
    // write('To-recipients of the item:');
            resolve(asyncResult);
        }
    }); // End getAsync for to-recipients.

  // Get any cc-recipients.
    ccRecipients.getAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    // write(asyncResult.error.message);
        } else {
    // Async call to get cc-recipients of the item completed.
    // Display the email addresses of the cc-recipients.
    // write('Cc-recipients of the item:');
            resolve(asyncResult);
        }
    }); // End getAsync for cc-recipients.

  // If the item has the bcc field, i.e., item is message,
  // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      // write(asyncResult.error.message);
            } else {
      // Async call to get bcc-recipients of the item completed.
      // Display the email addresses of the bcc-recipients.
      // write('Bcc-recipients of the item:');
                resolve(asyncResult);
            }
        }); // End getAsync for bcc-recipients.
    }
});


// Set the display name and email addresses of the recipients of
// the composed item.
OutlookUtil.setRecipients = () => {
  // Local objects to point to recipients of either
  // the appointment or message that is being composed.
  // bccRecipients applies to only messages, not appointments.
    let toRecipients,
        ccRecipients,
        bccRecipients;

  // Verify if the composed item is an appointment or message.
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    } else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }

  // Use asynchronous method setAsync to set each type of recipients
  // of the composed item. Each time, this example passes a set of
  // names and email addresses to set, and an anonymous
  // callback function that doesn't take any parameters.
    toRecipients.setAsync(
        [{
            'displayName': 'Graham Durkin',
            'emailAddress': 'graham@contoso.com'
        },
        {
            'displayName': 'Donnie Weinberg',
            'emailAddress': 'donnie@contoso.com'
        }],
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            // write(asyncResult.error.message);
            } else {
            // Async call to set to-recipients of the item completed.
            }
        }); // End to setAsync.


  // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
            'displayName': 'Perry Horning',
            'emailAddress': 'perry@contoso.com'
        },
        {
            'displayName': 'Guy Montenegro',
            'emailAddress': 'guy@contoso.com'
        }],
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            // write(asyncResult.error.message);
            } else {
            // Async call to set cc-recipients of the item completed.
            }
        }); // End cc setAsync.


  // If the item has the bcc field, i.e., item is message,
  // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                'displayName': 'Lewis Cate',
                'emailAddress': 'lewis@contoso.com'
            },
            {
                'displayName': 'Francisco Stitt',
                'emailAddress': 'francisco@contoso.com'
            }],
            (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                // write(asyncResult.error.message);
                } else {
                // Async call to set bcc-recipients of the item completed.
                // Do whatever appropriate for your scenario.
                }
            }); // End bcc setAsync.
    }
};


// Add specified recipients as required attendees of
// the composed appointment.
OutlookUtil.addAttendees = () => {
    if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
            [{
                'displayName': 'Kristie Jensen',
                'emailAddress': 'kristie@contoso.com'
            },
            {
                'displayName': 'Pansy Valenzuela',
                'emailAddress': 'pansy@contoso.com'
            }],
        (asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    // write(asyncResult.error.message);
            } else {
        // Async call to add attendees completed.
        // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
};


// Get the body type of the composed item, and prepend data
// in the appropriate data type in the item body.
OutlookUtil.prependItemBody = () => {
    item.body.getTypeAsync(
        (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                // write(asyncResult.error.message);
            }
            // Successfully got the type of item body.
            // Prepend data of the appropriate type in body.
            if (result.value === Office.MailboxEnums.BodyType.Html) {
            // Body is of HTML type.
            // Specify HTML in the coercionType parameter
            // of prependAsync.
                item.body.prependAsync(
                        '<b>Greetings!</b>',
                    { coercionType: Office.CoercionType.Html,
                        asyncContext: { var3: 1, var4: 2 } },
                        (asyncResult) => {
                            if (asyncResult.status ===
                                Office.AsyncResultStatus.Failed) {
                    // write(asyncResult.error.message);
                            } else {
                        // Successfully prepended data in item body.
                        // Do whatever appropriate for your scenario,
                        // using the arguments var3 and var4 as applicable.
                            }
                        });
            } else {
                // Body is of text type.
                item.body.prependAsync(
                        'Greetings!',
                    { coercionType: Office.CoercionType.Text,
                        asyncContext: { var3: 1, var4: 2 } },
                        (asyncResult) => {
                            if (asyncResult.status ===
                                Office.AsyncResultStatus.Failed) {
                    // write(asyncResult.error.message);
                            } else {
                        // Successfully prepended data in item body.
                        // Do whatever appropriate for your scenario,
                        // using the arguments var3 and var4 as applicable.
                            }
                        });
            }
        }
    );
};


// binding the office init
Office.initialize = () => {
    item = Office.context.mailbox.item;
    _resolvePromiseDoneInit();
};

export default OutlookUtil;
