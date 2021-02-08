# POPOAuthSample

This is a sample console application written in .Net Core that demonstrates how to obtain an OAuth token for logging on to a mailbox using POP.  Note that POP is a public protocol and as such it is up to the developer to correctly implement it in their code.  The example here is simplistic, and only intended to show how OAuth fits in to the log-in process.  You do not have to use MSAL to obtain the token, but it is a very simple way to do so.

You must register the application in Azure AD as per [this guide](https://docs.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth#get-an-access-token "Authenticate an IMAP application using OAuth").  You must add a redirect URL of http://localhost.

Once the application is registered, the application can be run from a command prompt (or PowerShell console).  The syntax is:

`POPOAuthSample TenantId ApplicationId`

If the parameters are valid, you will be prompted to log-in to the mailbox using the default system browser (POP only supports delegate access).  Once done, the application will use the token to log on to the mailbox and get a list of messages.  The POP conversation will be shown in the console.

A successful test looks like this:

![POPOAuthSample Successful Test Screenshot](https://github.com/David-Barrett-MS/POPOAuthSample/blob/master/POPOAuthSample.png?raw=true "POPOAuthSample Successful Test Screenshot")
