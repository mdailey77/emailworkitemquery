import { AzureFunction, Context } from "@azure/functions";
import * as azdev from "azure-devops-node-api";
import * as witapi from "azure-devops-node-api/WorkItemTrackingApi"

const timerTrigger: AzureFunction = async function (context: Context, myTimer: any): Promise<void> {
    let orgUrl = "https://dev.azure.com/{organization_name}"
    let token = process.env["DEVOPS_TOKEN"];
    let authHandler = azdev.getPersonalAccessTokenHandler(token); 
    let connection = new azdev.WebApi(orgUrl, authHandler);
    const ExcelJS = require('exceljs');
    const { DateTime } = require("luxon");
    const nodemailer = require("nodemailer");

    const queryId = "{work_item_query_id}";
    const wit: witapi.IWorkItemTrackingApi = await connection.getWorkItemTrackingApi();
    const result = await wit.queryById(queryId);

    if (result == null)
    {
        context.log("Result was null.");
    }

    // generate Excel file
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Azure DevOps';
    workbook.lastModifiedBy = 'Azure DevOps';
    workbook.created = new Date();
    const sheet = workbook.addWorksheet('Work Item Query');

    let colNames = [];
    let colRefNames: string[] = [];
    let queryItems = [];
    

    for (let i = 0; i < result.columns.length; i++) {
        let colObj = {header: '', key: '', width: 25, style: { font: { name: 'Roboto', bold: true } }};
        colObj.header = result.columns[i].name
        colObj.key = result.columns[i].referenceName;
        colNames[i] = colObj;
        colRefNames.push(result.columns[i].referenceName);
    }
    sheet.columns = colNames;

    for (let item of result.workItems) {
       queryItems.push(await wit.getWorkItem(item.id, colRefNames) )
    }

    queryItems.forEach((item) => {
        let sortingfields = {};
        let sortedfields;
        let rowArray = [];
        function luxonDate(datetimestring: string){
            return DateTime.fromISO(datetimestring).setZone('America/New_York').toLocaleString(DateTime.DATETIME_FULL);
        }
        Object.entries(item.fields).forEach(([key, value]) => {
            
            switch(key){
                case 'System.AssignedTo':
                    sortingfields[key] = item.fields[key].displayName;
                    break;
                case 'System.ChangedDate':
                    sortingfields[key] = luxonDate(item.fields[key]);
                    break;
                case 'Microsoft.VSTS.Common.ActivatedDate':
                    sortingfields[key] = luxonDate(item.fields[key]);
                    break;
                case 'Microsoft.VSTS.Common.ClosedDate':
                    sortingfields[key] = luxonDate(item.fields[key]);
                    break;
                default:
                    sortingfields[key] = item.fields[key];
                    break;
            }
        })


        const orderedData = (data, order) => order.map(k => data[k]);

        sortedfields = orderedData(sortingfields, colRefNames)

        rowArray = Object.values(sortedfields);


        let row = sheet.addRow(rowArray);
        row.font = { name: 'Roboto', size: 11, bold: false };
    })

    const currentdate = DateTime.now().toFormat('MM-dd-yy');
    context.log('creating Excel file');
    await workbook.xlsx.writeFile('C:\\local\\Temp\\workitemquery_' + currentdate + '.xlsx').then(() => {
        context.log('saved');
    }).catch((err) => {
        context.log('err', err);
    })

    // Send the email
    const transporter = nodemailer.createTransport({
        host: "{SMTP_email_server_IP}",
        port: 50025,
        secure: false,
        ignoreTLS: true,
        connectionTimeout: 180000,
        disableFileAccess: false
    });

    function setemailconfig() {
        const currentdatetime = DateTime.now().setZone('America/New_York').toLocaleString(DateTime.DATETIME_FULL);
        const from = 'NO_REPLY@website.com';
        const to = 'recipent@website.com'
        const cc = 'anotherrecipent@website.com'
        const bcc = 'thirdrecipent@website.com'
        const subject = 'Work Item Query ' + currentdate;
        const text = 'Here is the latest work item query';
        const html = '<p>Here is the latest work item query that was run on ' + currentdatetime + '</p><br>Azure DevOps';
        
        const attachments = [
            {
                filename: 'workitemquery_' + currentdate + '.xlsx',
                path: 'C:\\local\\Temp\\workitemquery_' + currentdate + '.xlsx'
            }
        ];

        var mailOption = {
            from: from,
            to:  to,
            cc: cc,
            bcc: bcc,
            subject: subject,
            text: text,
            html: html,
            attachments: attachments
        }
        return mailOption;
    }

    context.log('Sending email');
    let info = await transporter.sendMail(setemailconfig(), function(error, info) {
            if (error) {
                context.log('sent mail error: ' + error.message);
            }
            context.log('sent mail successful: ' + info.response);
        }
    );
    
    context.log('sendMail function ' + info);
    
};

export default timerTrigger;
