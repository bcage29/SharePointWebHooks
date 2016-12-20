// Web Hook
module.exports = function (context, data) {
    context.log("logging request object", context.req);
    context.log("logging data object", data);

    if (context.req) {
        if (context.req.query && context.req.query.validationtoken) {
            context.log("validating the web hook");
            // We need to return the validation token immediately in order
            // for the WebHook to register
            context.log("validationToken = " + context.req.query.validationtoken);
            context.res = {
                status: 200,
                // Default content type is json
                // SharePoint requires text/plain
                Headers: {
                    "Content-Type": "text/plain"
                },
                body: context.req.query.validationtoken
            };
        } else {
            try {
                context.log("begin processing the new event(s)");

                // initialize reference to updateItem. anything added to this array will be 
                // processed to our outbound queue binding
                context.bindings.updateItem = [];
                context.req.body.value.forEach((singleEvent) => {

                    // validate inbound payload
                    if (!singleEvent.subscriptionId) throw 'Subscription ID required';
                    if (!singleEvent.clientState) throw 'Client state required';
                    if (!singleEvent.expirationDateTime) throw 'Expiration date required';
                    if (!singleEvent.resource) throw 'Resource ID required';
                    if (!singleEvent.tenantId) throw 'Tenant IT required';
                    if (!singleEvent.siteUrl) throw 'Site URL required';
                    if (!singleEvent.webId) throw 'Web ID required';

                    context.bindings.updateItem.push(singleEvent);
                    context.log("added item to queue");
                });

                context.log('Finished adding items to the queue');

                // return success response
                context.res = { status: 200 };
            } catch (e) {
                // return error response with detailed message
                context.log('An error occured' + e);
                context.res = {
                    status: 400,
                    body: { error: e }
                };
            }
        }
    }
    context.done();
};