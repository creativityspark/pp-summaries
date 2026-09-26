/**
 * One-off diagnostic: try creating csp_AIUsage.csp_Prompt lookup → csp_aiprompt
 * and print the FULL error body from Dataverse (instead of the truncated message
 * that create-tables.mjs prints).
 */

import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const API_VERSION = "v9.2";

const pca = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
});

const token = (
  await pca.acquireTokenByDeviceCode({
    scopes: [`${ENV_URL}/.default`],
    deviceCodeCallback: (r) => {
      console.log("\n=========================================");
      console.log("DEVICE CODE AUTHENTICATION");
      console.log("=========================================");
      console.log(r.message);
      console.log("=========================================\n");
    },
  })
).accessToken;
console.log("Authenticated.\n");

const label = (text) => ({
  LocalizedLabels: [{ Label: text, LanguageCode: 1033 }],
});

const body = {
  "@odata.type": "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata",
  SchemaName: "csp_csp_aiprompt_csp_aiusage_Prompt",
  ReferencedEntity: "csp_aiprompt",
  ReferencingEntity: "csp_aiusage",
  ReferencedAttribute: "csp_aipromptid",
  Lookup: {
    "@odata.type": "Microsoft.Dynamics.CRM.LookupAttributeMetadata",
    AttributeType: "Lookup",
    AttributeTypeName: { Value: "LookupType" },
    SchemaName: "csp_Prompt",
    DisplayName: label("Prompt"),
    Description: label(
      "Prompt that was invoked. Optional so historical rows survive even if the prompt is later deleted; csp_PromptName captures the name as a fallback."
    ),
    RequiredLevel: { Value: "None" },
  },
  AssociatedMenuConfiguration: {
    Behavior: "UseCollectionName",
    Group: "Details",
    Order: 10000,
    IsCustomizable: true,
  },
  CascadeConfiguration: {
    Assign: "NoCascade",
    Delete: "RemoveLink",
    Merge: "NoCascade",
    Reparent: "NoCascade",
    Share: "NoCascade",
    Unshare: "NoCascade",
  },
};

console.log("POST /RelationshipDefinitions with body:");
console.log(JSON.stringify(body, null, 2));
console.log("");

const res = await fetch(
  `${ENV_URL}/api/data/${API_VERSION}/RelationshipDefinitions?SolutionUniqueName=BizzSummit2026`,
  {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "MSCRM.SolutionUniqueName": "BizzSummit2026",
    },
    body: JSON.stringify(body),
  }
);

console.log(`Response: ${res.status} ${res.statusText}\n`);
const text = await res.text();
console.log("Body:");
console.log(text);
