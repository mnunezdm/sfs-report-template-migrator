/*
 * Copyright (c) 2022, Daniel Nakonieczny
 * All rights reserved.
 * date: June 28 2022
 * description: Contains the Service Report Template Migrator code
 */

import { launch } from 'puppeteer';
import { parse } from 'yaml';
import jsforce from 'jsforce';
import { appendFile, writeFile, readFile } from "fs/promises";
import { config } from 'dotenv';
import yargs from 'yargs';
import { hideBin } from 'yargs/helpers';
config();

const argv = yargs(hideBin(process.argv))
  .command("retrieve")
  .command("deploy")
  .command("all")
  .option('headless', {
    alias: 'x',
    boolean: true,
    default: false
  })
  .strictCommands()
  .strictOptions()
  .demandCommand(1, 2, "Please select a command", "Not more than 1 command available")
  .argv;

const CUSTOM_FIELD_ID_REGEX = /00N[a-zA-Z0-9]{12}/g;
const JSON_LAYOUT_PARAM_REGEX = /j_id0%3Af%3AjsonLayout[^&?]*?=[^&?]*/;
const IMG_TAG_REGEX_1 = /(?<=%3Cimg).*?(?=%2F%3E)/g;
const IMG_TAG_REGEX_2 = /(?<=%3Cimg).*?(?=%3C%2Fimg%3E)/g;
const EMPTY_IMG_TAG_1 = '%3Cimg%2F%3E';
const EMPTY_IMG_TAG_2 = '%3Cimg%3C%2Fimg%3E';
const reportNamesFile = await readFile('./config.yml', 'utf8');
const yamlConfig = parse(reportNamesFile);
const reportNames = yamlConfig.reportNames;
const subtypesToMigrate = yamlConfig.reportSubtypesToMigrate;
const LOG_POST_DATA = yamlConfig.writePOSTDataToFile;
const ERROR_LOG_FILENAME = yamlConfig.errorLogFilename;
const WINDOW_WIDTH = yamlConfig.windowWidth;
const WINDOW_HEIGHT = yamlConfig.windowHeight;
const TIMEOUT_BETWEEN_ACTIONS = yamlConfig.timeoutBetweenActions;
const REPLACE_SOURCE_IMAGES = yamlConfig.removeSourceImages;
const IMAGE_REPLACEMENT_TEXT = yamlConfig.imageReplacementText;
const SUPPORTED_SUBTYPES = {
  SA_WO: 'Service Appointment for Work Order',
  SA_WOLI: 'Service Appointment for Work Order Line Item',
  WO: 'Work Order',
  WOLI: 'Work Order Line Item',
};

let browser;
let incognitoContext;
let openedPages = [];
let reportNameToJSON = {};
let reportNameToJSONReplaced = {};
let customObjectIdsToLookup = [];
let sourceObjectIdToNameMap = {};
let sourceObjectIdToSObject = {};
let sourceObjectIdMap = {};
let api_name_list = [];

function sleep(ms) {
  return new Promise(resolve => {
    setTimeout(resolve, ms);
  });
}

const logAndExit = async stringError => {
  try {
    await appendFile(ERROR_LOG_FILENAME, stringError);
  } catch (err) {
    console.error(err);
  }
  throw new Error(`\r\n${stringError}\r\n`);
};

await writeFile(ERROR_LOG_FILENAME, '', { flag: 'a' });

console.log('info: login to source environment');
const sourceConnection = await loginEnvironment(
  process.env.SOURCE_ORG_LOGIN_URL,
  process.env.SOURCE_ORG_ACCESS_TOKEN,
  process.env.SOURCE_ORG_USERNAME,
  process.env.SOURCE_ORG_PASSWORD,
  process.env.SOURCE_ORG_SECURITY_TOKEN,
)

console.log('info: login to target environment');
const targetConnection = await loginEnvironment(
  process.env.TARGET_ORG_LOGIN_URL,
  process.env.TARGET_ORG_ACCESS_TOKEN,
  process.env.TARGET_ORG_USERNAME,
  process.env.TARGET_ORG_PASSWORD,
  process.env.TARGET_ORG_SECURITY_TOKEN,
)

await validateReports(reportNames)

console.log(`info: opening browser ${argv.headless && 'in headless mode' || ''}`);
browser = await launch({
  headless: argv.headless,
  args: [`--window-size=${WINDOW_WIDTH},${WINDOW_HEIGHT}`],
  defaultViewport: {
    width: WINDOW_WIDTH,
    height: WINDOW_HEIGHT,
  },
});

incognitoContext = await browser.createIncognitoBrowserContext();
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: Login to source in browser');
await loginToOrg(sourceConnection.loginUrl, sourceConnection.accessToken);
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: Login to target in browser');
await loginToOrg(targetConnection.loginUrl, targetConnection.accessToken, true);
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: grabing source report links');
const reportNameToURLMapSource = await grabSourceOrgReportLinks();
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: Grabing source reports');
await grabSourceOrgReportJSON(reportNameToURLMapSource);
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: Cleaning tabs');
await cleanupTabs();
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: extracting source org information')
await extractSourceOrgCustomObjectsAndFields();
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: matching to target org information')
const sourceFieldIdMap = await matchSourceToTargetOrgCustomObjectAndFieldIds();
await sleep(TIMEOUT_BETWEEN_ACTIONS);
console.log('info: replacing information in target org')
await replaceCustomFieldIds(sourceFieldIdMap);
console.log('info: Creating missing reports in target');
await createReportsInTargetOrg(reportNames);
console.log('info: Grabing target report links');
await sleep(TIMEOUT_BETWEEN_ACTIONS);
const reportNameToURLMapTarget = await grabTargetOrgReportLinks();
await sleep(TIMEOUT_BETWEEN_ACTIONS);
await deployReportTemplatesToTargetOrg(reportNameToURLMapTarget);

console.log('info: success!');

await browser.close();

async function loginEnvironment(loginUrl, accessToken, username, password, token) {
  const connection = new jsforce.Connection({
    loginUrl,
    accessToken,
    instanceUrl: loginUrl,
    version: '55.0',
  });
  if (!connection.accessToken) {
    await sourceConnection.login(
      username,
      `${password}${token}`,
    );
  }
  return connection;
}

/**
 * 
 * @param {String[]} reportNames 
 */
async function validateReports(reportNames) {
  console.log('fine: validating reports in source org')
  const response = await sourceConnection.query(`SELECT Id, DeveloperName FROM ServiceReportLayout WHERE DeveloperName IN ('${reportNames.join("','")}')`)
  const reportsInOrg = new Set(response.records.map(report => report.DeveloperName));
  const missingReports = reportNames
    .filter(report => !reportsInOrg.has(report))

  if (missingReports.length) {
    throw new Error(`Missing reports ${missingReports.join(', ')}`)
  }
}

async function logErrors(messagesArray) {
  console.error('The following errors happened:');
  console.error(messagesArray);
  let formattedErrors = messagesArray.map(
    message => `${new Date().toLocaleString()} ${message}`,
  );
  logAndExit(`${formattedErrors.join('\r\n')}\r\n`);
}

async function loginToOrg(loginUrl, accessToken, incognito) {
  if (!accessToken) {
    let message = 'Browser login failed. Please run this script again.';
    logAndExit(`${new Date().toLocaleString()} ${message}\r\n`);
  }

  let loginPage;

  if (incognito) {
    loginPage = await incognitoContext.newPage();
  } else {
    loginPage = await browser.newPage();
  }

  await loginPage.goto(`${loginUrl}/secur/frontdoor.jsp?sid=${accessToken}`);
  await loginPage.waitForTimeout(TIMEOUT_BETWEEN_ACTIONS);

  const pageUrl = await loginPage.url();

  if (pageUrl.includes('ec=302')) {
    //sometimes the frontdoor.jsp login doesn't work and the script needs to be restarted
    let message = 'Browser login failed. Please run this script again.';
    logAndExit(`${new Date().toLocaleString()} ${message}\r\n`);
  }
}

async function cleanupTabs() {
  for (let p of openedPages) {
    p.close();
  }

  openedPages = [];
}

async function createReportsInTargetOrg() {
  const query = `SELECT Id, DeveloperName, MasterLabel, TemplateType FROM ServiceReportLayout WHERE DeveloperName IN ('${reportNames.join("','")}')`;
  const response = await targetConnection.query(query)
  const reportsInOrg = new Set(response.records.map(report => report.DeveloperName));
  const missingReports = reportNames
    .filter(report => !reportsInOrg.has(report))

  if (!missingReports.length) {
    console.log('fine: all records exist in target org, doing nothing...')
  }

  for (let reportName of missingReports) {
    console.log(`fine: '${reportName}' does not exist or is inactive, creating it`)
    let newReportPage = await incognitoContext.newPage();
    await newReportPage.goto(
      `${process.env.TARGET_ORG_LOGIN_URL}/_ui/support/fieldservice/ui/ServiceReportTemplateClone/e?p1=${reportName}`,
      { waitUntil: 'networkidle0' },
    );
    await newReportPage.click("input[name='save']");
    await sleep(TIMEOUT_BETWEEN_ACTIONS);
  }
}

async function grabSourceOrgReportLinks() {
  const reportNameToURLMapSource = {};
  for (const reportName of reportNames) {
    console.log(`fine: getting report link for ${reportName}`);
    let newReportPage = await browser.newPage();
    await newReportPage.goto(
      `${process.env.SOURCE_ORG_LOGIN_URL}/_ui/support/fieldservice/ui/ServiceReportTemplateLayouts`,
    );
    const reportLink = await newReportPage.evaluate(
      name =>
        document.querySelector(`a[title$="${name}"]`).getAttribute('href'),
      reportName,
    );
    reportNameToURLMapSource[reportName] = reportLink;
  }
  return reportNameToURLMapSource;
}

async function grabSourceOrgReportJSON(reportNameToURLMapSource) {
  for (let reportSubtype of Object.keys(SUPPORTED_SUBTYPES)) {
    if (subtypesToMigrate.includes(reportSubtype)) {
      await grabSourceReport(
        reportSubtype,
        SUPPORTED_SUBTYPES[reportSubtype],
        reportNameToURLMapSource,
      )
    }
  }
}

async function grabSourceReport(
  subtypeName,
  subtypeLabel,
  reportNameToURLMapSource,
) {
  const requestsProcessed = []
  console.log(`fine: getting reports for subtype '${subtypeLabel}'`);
  for (const currentReportName in reportNameToURLMapSource) {
    console.log(
      `fine: getting report for '${currentReportName}' for subtype '${subtypeLabel}'`,
    );
    const reportVersionName = `${currentReportName}_${subtypeName}`;
    const url = reportNameToURLMapSource[currentReportName];

    let newReportPage = await browser.newPage();
    openedPages.push(newReportPage);

    await newReportPage.goto(`${process.env.SOURCE_ORG_LOGIN_URL}${url}`, {
      waitUntil: 'networkidle0',
    });
    await goToTemplateSubtype(newReportPage, subtypeLabel);

    await newReportPage.setRequestInterception(true);

    newReportPage.on('request', request => {
      const request_url = request.url();
      const request_post_data = request.postData();

      if (
        request_url.includes(
          '/servicereport/serviceReportTemplateEditor.apexp',
        ) &&
        request_post_data &&
        request_post_data.includes('j_id0%3Af%3AjsonLayout') &&
        !requestsProcessed.includes(reportVersionName)
      ) {
        const regex = JSON_LAYOUT_PARAM_REGEX;
        const matched = regex.exec(request_post_data);
        reportNameToJSON[reportVersionName] = matched[0];

        if (LOG_POST_DATA) {
          let dataToWrite = decodeURIComponent(
            matched[0].replace('j_id0%3Af%3AjsonLayout=', ''),
          );
          const dataToWriteFormatted = JSON.stringify(
            JSON.parse(dataToWrite),
            null,
            2,
          );

          writeFile(
            `${reportVersionName}.source.json`,
            dataToWriteFormatted,
          ).catch(err => {
            console.error(err);
          });
        }

        requestsProcessed.push(reportVersionName);

        request.continue();
      } else {
        request.continue();
      }
    });

    await clickQuickSave(newReportPage);
    return requestsProcessed;
  }
}

async function extractSourceOrgCustomObjectsAndFields() {
  const fieldIds = new Set()
  for (const currentReportName in reportNameToJSON) {
    const jsonString = reportNameToJSON[currentReportName];
    const array = [...jsonString.matchAll(CUSTOM_FIELD_ID_REGEX)];
    const results = array.flatMap(x => x[0]);

    for (const fieldId of results) {
      fieldIds.add(fieldId)
    }
  }

  const response = await sourceConnection.tooling.query(
    `SELECT Id, DeveloperName, NamespacePrefix, TableEnumOrId FROM CustomField WHERE Id IN ('${[...fieldIds].join("', '")}')`,
  );
  response.records.forEach(record => {
    if (response.records) {
      api_name_list.push(record);

      if (
        record.TableEnumOrId.startsWith('01I') &&
        !customObjectIdsToLookup.includes(record.TableEnumOrId)
      ) {
        console.log(`adding source TableEnumOrId: ${record.TableEnumOrId}`);
        customObjectIdsToLookup.push(record.TableEnumOrId);
      }
    }
  });

  console.log(`info: looking for customObjectIdsToLookup: ${customObjectIdsToLookup}`);
  if (customObjectIdsToLookup.length) {
    const responseCustomObject = await sourceConnection.tooling.query(
      `SELECT Id, DeveloperName, NamespacePrefix FROM CustomObject WHERE Id IN ('${customObjectId.join("','")}')`
    )
    responseCustomObject.records.forEach((record) => {
      console.log(`custom object: ${record.DeveloperName}`);
      sourceObjectIdToNameMap[customObjectId] = record;
    })
  }
}

async function matchSourceToTargetOrgCustomObjectAndFieldIds() {
  let missingObjects = [];

  for (const sourceObjectId in sourceObjectIdToNameMap) {
    const customObject = sourceObjectIdToNameMap[sourceObjectId];

    if (customObject && customObject.DeveloperName) {
      const query = `SELECT Id, DeveloperName, NamespacePrefix FROM CustomObject WHERE DeveloperName = '${
        customObject.DeveloperName
      }' AND NamespacePrefix = '${
        customObject.NamespacePrefix == null
          ? ''
          : customObject.NamespacePrefix
      }'`
      const response = await targetConnection.tooling.query(query);
      if (response.records && res.records[0]) {
        let record = res.records[0];
        sourceObjectIdMap[customObject.Id] = record.Id;
        sourceObjectIdToSObject[customObject.Id] = record;
        console.log(
          `adding object source id: ${customObject.Id}, target id: ${record.Id}`,
        );
      } else {
        let errorMessage = `custom object missing in target org: ${
          customObject.NamespacePrefix == null
            ? ''
            : customObject.NamespacePrefix + '__'
        }${customObject.DeveloperName}__c`;
        missingObjects.push(errorMessage);
      }
    }
  }

  if (missingObjects.length) {
    logErrors(missingObjects);
  }
  
  const getFieldApiName = (field) => {
    const ns = field.NamespacePrefix ? field.NamespacePrefix : '';
    const table = field.TableEnumOrId.startsWith('01I') ? 
      sourceObjectIdMap[field.TableEnumOrId] :
      field.TableEnumOrId;
    return `${table}.${ns}${field.DeveloperName}`;
  }

  const fieldFilters = api_name_list
    .map(field => {
      const ns = field.NamespacePrefix ? field.NamespacePrefix : '';
      const dn = field.DeveloperName;
      const table = field.TableEnumOrId.startsWith('01I') ? 
        sourceObjectIdMap[field.TableEnumOrId] :
        field.TableEnumOrId;
      return `DeveloperName = '${dn}' AND TableEnumOrId = '${table}' AND NamespacePrefix = '${ns}'`
    });
  const fieldsByApi = Object.fromEntries(
    api_name_list.map(field => [getFieldApiName(field), field])
  );

  const query2 = `SELECT Id, DeveloperName, NamespacePrefix, TableEnumOrId FROM CustomField WHERE (${
    fieldFilters.join(') OR (')
  })`

  const response2 = await targetConnection.tooling.query(query2);
  
  const missingFields = response2.records
    .filter(field => fieldsByApi[getFieldApiName(field)])
    .keys();
    
  const sourceFieldIdMap = Object.fromEntries(response2.records
    .filter(field => fieldsByApi[getFieldApiName(field)])  
    .map(targetField => 
      [fieldsByApi[getFieldApiName(targetField)].Id.substring(0, 15), targetField.Id.substring(0, 15)]
    ))

  if (missingFields.length) {
    logErrors(`Missing fields in target org ${missingFields}`);
  }
  
  return sourceFieldIdMap;
}

async function replaceCustomFieldIds(sourceFieldIdMap) {
  for (const currentReportName in reportNameToJSON) {
    let jsonString = reportNameToJSON[currentReportName];
    const array = [...jsonString.matchAll(CUSTOM_FIELD_ID_REGEX)];
    const results = array.flatMap(x => x[0]);
    console.log(
      `fine: custom field Ids found in source org for ${currentReportName}:`,
    );

    for (const fieldId of results) {
      if (sourceFieldIdMap[fieldId]) {
        const targetOrgFieldId = sourceFieldIdMap[fieldId];
        console.log(
          `fine: target org Id of source custom field ${fieldId}: ${targetOrgFieldId}`,
        );
        jsonString = jsonString.replaceAll(fieldId, targetOrgFieldId);
      }
    }

    if (REPLACE_SOURCE_IMAGES) {
      jsonString = jsonString.replaceAll(IMG_TAG_REGEX_1, '');
      jsonString = jsonString.replaceAll(IMG_TAG_REGEX_2, '');
      jsonString = jsonString.replaceAll(EMPTY_IMG_TAG_1, IMAGE_REPLACEMENT_TEXT);
      jsonString = jsonString.replaceAll(EMPTY_IMG_TAG_2, IMAGE_REPLACEMENT_TEXT);
    }

    reportNameToJSONReplaced[currentReportName] = jsonString;
  }
}

async function grabTargetOrgReportLinks() {
  const reportNameToURLMapTarget = {};
  for (const reportName of reportNames) {
    let newReportPage = await incognitoContext.newPage();
    await newReportPage.goto(
      `${process.env.TARGET_ORG_LOGIN_URL}/_ui/support/fieldservice/ui/ServiceReportTemplateLayouts`,
    );
    const reportLink = await newReportPage.evaluate(
      name =>
        document.querySelector(`a[title$="${name}"]`).getAttribute('href'),
      reportName,
    );
    reportNameToURLMapTarget[reportName] = reportLink;
  }
  return reportNameToURLMapTarget;
}

async function deployReportTemplatesToTargetOrg(reportNameToURLMapTarget) {
  let requestsProcessed = [];

  for (let reportSubtype of Object.keys(SUPPORTED_SUBTYPES)) {
    if (subtypesToMigrate.includes(reportSubtype)) {
      requestsProcessed = [
        ...requestsProcessed,
        ...await deployReportTemplate(
          reportSubtype,
          SUPPORTED_SUBTYPES[reportSubtype],
          reportNameToURLMapTarget,
        )
      ]
    }
  }
}

async function deployReportTemplate(
  subtypeName,
  subtypeLabel,
  reportNameToURLMapTarget,
) {
  const requestsProcessed = [];
  console.log(`fine: deploying reports for subtype '${subtypeLabel}'`);
  for (const currentReportName in reportNameToURLMapTarget) {
    console.log(
      `fine: deploying '${currentReportName}' report for subtype '${subtypeLabel}'`,
    );
    const reportVersionName = `${currentReportName}_${subtypeName}`;
    const url = reportNameToURLMapTarget[currentReportName];

    let newReportPage = await incognitoContext.newPage();

    openedPages.push(newReportPage);

    await newReportPage.goto(`${process.env.TARGET_ORG_LOGIN_URL}${url}`, {
      waitUntil: 'networkidle0',
    });
    await goToTemplateSubtype(newReportPage, subtypeLabel);

    await newReportPage.setRequestInterception(true);

    newReportPage.on('request', request => {
      const request_url = request.url();
      const request_post_data = request.postData();

      if (
        request_url.includes(
          '/servicereport/serviceReportTemplateEditor.apexp',
        ) &&
        request_post_data?.includes('j_id0%3Af%3AjsonLayout')
      ) {
        const regex = JSON_LAYOUT_PARAM_REGEX;
        const matchedString = regex.exec(request_post_data)[0];

        if (reportNameToJSON[reportVersionName]) {
          request.continue({
            postData: request_post_data.replace(
              matchedString,
              reportNameToJSONReplaced[reportVersionName],
            ),
          });

          if (LOG_POST_DATA && !requestsProcessed.includes(reportVersionName)) {
            let dataToWrite = decodeURIComponent(
              reportNameToJSONReplaced[reportVersionName].replace(
                'j_id0%3Af%3AjsonLayout=',
                '',
              ),
            );
            const dataToWriteFormatted = JSON.stringify(
              JSON.parse(dataToWrite),
              null,
              2,
            );

            writeFile(
              `${reportVersionName}.target.json`,
              dataToWriteFormatted,
            ).catch(err => {
              console.error(err);
            });
          }

          requestsProcessed.push(reportVersionName);
        } else {
          request.continue();
        }
      } else {
        request.continue();
      }
    });

    await clickQuickSave(newReportPage);
    return requestsProcessed;
  }
}

async function clickQuickSave(reportPage) {
  const [button] = await reportPage.$x("//button[contains(., 'Quick Save')]");
  if (button) {
    await button.click();
    await reportPage.waitForTimeout(TIMEOUT_BETWEEN_ACTIONS);
  }

  await sleep(TIMEOUT_BETWEEN_ACTIONS);
}

async function goToTemplateSubtype(reportPage, subtypeLabel) {
  await reportPage.waitForTimeout(TIMEOUT_BETWEEN_ACTIONS);
  let optionValue = await reportPage.$$eval(
    'select[name$="childLayoutPicklist:templateList"] option',
    (options, subtypeLabel) =>
      options.find(o => o.innerText === subtypeLabel)?.value,
    subtypeLabel,
  );
  await reportPage.select(
    'select[name$="childLayoutPicklist:templateList"]',
    optionValue,
  );
  await reportPage.waitForTimeout(TIMEOUT_BETWEEN_ACTIONS);
}
