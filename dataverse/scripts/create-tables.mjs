import { PublicClientApplication } from "@azure/msal-node";
import { readFile } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import path from "node:path";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const API_VERSION = "v9.2";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const label = (text) => ({
  LocalizedLabels: [{ Label: text, LanguageCode: 1033 }]
});

async function getToken() {
  const pca = new PublicClientApplication({
    auth: {
      clientId: CLIENT_ID,
      authority: `https://login.microsoftonline.com/${TENANT_ID}`
    }
  });
  const result = await pca.acquireTokenByDeviceCode({
    scopes: [`${ENV_URL}/.default`],
    deviceCodeCallback: (response) => {
      console.log("\n=========================================");
      console.log("DEVICE CODE AUTHENTICATION");
      console.log("=========================================");
      console.log(response.message);
      console.log("=========================================\n");
    }
  });
  return result.accessToken;
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function api(token, method, urlPath, body, extraHeaders = {}) {
  const url = `${ENV_URL}/api/data/${API_VERSION}${urlPath}`;
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
      "MSCRM.SolutionUniqueName": "BizzSummit2026",
      ...extraHeaders
    },
    body: body ? JSON.stringify(body) : undefined
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`${method} ${urlPath} -> ${res.status} ${res.statusText}\n${text}`);
  }
  if (res.status === 204) return null;
  const ct = res.headers.get("content-type") || "";
  return ct.includes("application/json") ? res.json() : res.text();
}

async function apiWithRetry(token, method, urlPath, body, extraHeaders = {}) {
  const maxAttempts = 4;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const result = await api(token, method, urlPath, body, extraHeaders);
      return result;
    } catch (err) {
      const transient = /\b(400|429|500|502|503|504)\b/.test(err.message);
      if (attempt < maxAttempts && transient) {
        const wait = 1500 * attempt;
        console.log(`    [retry ${attempt}/${maxAttempts - 1}] after ${wait}ms - ${err.message.split("\n")[0]}`);
        await sleep(wait);
        continue;
      }
      throw err;
    }
  }
}

function buildPrimaryAttribute(primary) {
  return {
    "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
    AttributeType: "String",
    AttributeTypeName: { Value: "StringType" },
    SchemaName: primary.schemaName,
    RequiredLevel: { Value: primary.requiredLevel || "None" },
    DisplayName: label(primary.displayName),
    Description: label(primary.description),
    MaxLength: primary.maxLength ?? 100,
    FormatName: { Value: "Text" },
    IsPrimaryName: true
  };
}

function buildAttribute(prefix, attr) {
  const common = {
    SchemaName: attr.schemaName,
    DisplayName: label(attr.displayName),
    Description: label(attr.description),
    RequiredLevel: { Value: attr.requiredLevel || "None" }
  };

  switch (attr.type) {
    case "String":
      return {
        "@odata.type": "Microsoft.Dynamics.CRM.StringAttributeMetadata",
        AttributeType: "String",
        AttributeTypeName: { Value: "StringType" },
        MaxLength: attr.maxLength ?? 200,
        FormatName: { Value: attr.format || "Text" },
        ...common
      };
    case "Memo":
      return {
        "@odata.type": "Microsoft.Dynamics.CRM.MemoAttributeMetadata",
        AttributeType: "Memo",
        AttributeTypeName: { Value: "MemoType" },
        MaxLength: attr.maxLength ?? 4000,
        Format: "TextArea",
        ImeMode: "Disabled",
        ...common
      };
    case "Integer":
      return {
        "@odata.type": "Microsoft.Dynamics.CRM.IntegerAttributeMetadata",
        AttributeType: "Integer",
        AttributeTypeName: { Value: "IntegerType" },
        Format: "None",
        MinValue: attr.minValue ?? -2147483648,
        MaxValue: attr.maxValue ?? 2147483647,
        ...common
      };
    case "Decimal":
      return {
        "@odata.type": "Microsoft.Dynamics.CRM.DecimalAttributeMetadata",
        AttributeType: "Decimal",
        AttributeTypeName: { Value: "DecimalType" },
        Precision: attr.precision ?? 2,
        MinValue: attr.minValue ?? -100000000000,
        MaxValue: attr.maxValue ?? 100000000000,
        ...common
      };
    case "DateTime":
      return {
        "@odata.type": "Microsoft.Dynamics.CRM.DateTimeAttributeMetadata",
        AttributeType: "DateTime",
        AttributeTypeName: { Value: "DateTimeType" },
        Format: "DateAndTime",
        DateTimeBehavior: { Value: "UserLocal" },
        ImeMode: "Disabled",
        ...common
      };
    case "Boolean":
      return {
        "@odata.type": "Microsoft.Dynamics.CRM.BooleanAttributeMetadata",
        AttributeType: "Boolean",
        AttributeTypeName: { Value: "BooleanType" },
        DefaultValue: false,
        OptionSet: {
          TrueOption: { Value: 1, Label: label("Yes") },
          FalseOption: { Value: 0, Label: label("No") }
        },
        ...common
      };
    case "Picklist":
      return {
        "@odata.type": "Microsoft.Dynamics.CRM.PicklistAttributeMetadata",
        AttributeType: "Picklist",
        AttributeTypeName: { Value: "PicklistType" },
        OptionSet: {
          "@odata.type": "Microsoft.Dynamics.CRM.OptionSetMetadata",
          OptionSetType: "Picklist",
          IsGlobal: false,
          Options: attr.options.map((opt) => ({
            Value: opt.value,
            Label: label(opt.label)
          }))
        },
        ...common
      };
    default:
      throw new Error(`Unsupported attribute type: ${attr.type}`);
  }
}

function buildOneToManyRelationship(prefix, referencingEntity, attr) {
  const referencedEntity = attr.targetEntity;
  const referencedAttribute = referencedEntity === "systemuser"
    ? "systemuserid"
    : `${referencedEntity}id`;
  const schemaName = `${prefix}_${referencedEntity}_${referencingEntity}_${attr.schemaName.replace(/^csp_/, "")}`;
  return {
    "@odata.type": "Microsoft.Dynamics.CRM.OneToManyRelationshipMetadata",
    SchemaName: schemaName,
    ReferencedEntity: referencedEntity,
    ReferencingEntity: referencingEntity,
    ReferencedAttribute: referencedAttribute,
    Lookup: {
      "@odata.type": "Microsoft.Dynamics.CRM.LookupAttributeMetadata",
      AttributeType: "Lookup",
      AttributeTypeName: { Value: "LookupType" },
      SchemaName: attr.schemaName,
      DisplayName: label(attr.displayName),
      Description: label(attr.description),
      RequiredLevel: { Value: attr.requiredLevel || "None" }
    },
    AssociatedMenuConfiguration: {
      Behavior: "UseCollectionName",
      Group: "Details",
      Order: 10000,
      IsCustomizable: true
    },
    CascadeConfiguration: {
      Assign: "NoCascade",
      Delete: "RemoveLink",
      Merge: "NoCascade",
      Reparent: "NoCascade",
      Share: "NoCascade",
      Unshare: "NoCascade"
    }
  };
}

function referencedAttributeFor(entity) {
  if (entity === "systemuser") return "systemuserid";
  return `${entity}id`;
}

function buildPolymorphicLookup(prefix, referencingEntity, attr) {
  const attrShortName = attr.schemaName.replace(/^csp_/, "");
  const relationships = attr.targets.map((target) => ({
    "@odata.type": "Microsoft.Dynamics.CRM.ComplexOneToManyRelationshipMetadata",
    SchemaName: `${prefix}_${target}_${referencingEntity}_${attrShortName}`,
    ReferencedEntity: target,
    ReferencingEntity: referencingEntity,
    ReferencedAttribute: referencedAttributeFor(target)
  }));

  return {
    OneToManyRelationships: relationships,
    Lookup: {
      "@odata.type": "Microsoft.Dynamics.CRM.ComplexLookupAttributeMetadata",
      AttributeType: "Lookup",
      AttributeTypeName: { Value: "LookupType" },
      SchemaName: attr.schemaName,
      DisplayName: label(attr.displayName),
      Description: label(attr.description),
      RequiredLevel: { Value: attr.requiredLevel || "None" }
    },
    SolutionUniqueName: "BizzSummit2026"
  };
}

function buildEntityBody(prefix, table) {
  const logicalName = table.schemaName.toLowerCase();
  return {
    "@odata.type": "Microsoft.Dynamics.CRM.EntityMetadata",
    SchemaName: table.schemaName,
    LogicalName: logicalName,
    DisplayName: label(table.displayName),
    DisplayCollectionName: label(table.displayCollectionName),
    Description: label(table.description),
    HasActivities: false,
    HasNotes: table.hasNotes ?? false,
    IsActivity: false,
    OwnershipType: "UserOwned",
    Attributes: [buildPrimaryAttribute(table.primaryAttribute)],
    PrimaryNameAttribute: table.primaryAttribute.schemaName.toLowerCase()
  };
}

async function tableExists(token, logicalName) {
  try {
    await api(token, "GET", `/EntityDefinitions(LogicalName='${logicalName}')?$select=LogicalName`);
    return true;
  } catch (err) {
    if (err.message.includes("404")) return false;
    throw err;
  }
}

async function attributeExists(token, logicalName, attrLogicalName) {
  try {
    await api(token, "GET", `/EntityDefinitions(LogicalName='${logicalName}')/Attributes(LogicalName='${attrLogicalName}')?$select=LogicalName`);
    return true;
  } catch (err) {
    if (err.message.includes("404")) return false;
    throw err;
  }
}

async function main() {
  const tablesPath = path.join(__dirname, "tables.json");
  const config = JSON.parse(await readFile(tablesPath, "utf-8"));
  const { tables, prefix } = config;

  console.log(`Authenticating against tenant ${TENANT_ID}...`);
  const token = await getToken();
  console.log("Authenticated.\n");

  const failures = [];

  for (const table of tables) {
    const logicalName = table.schemaName.toLowerCase();
    console.log(`\n=== Table: ${table.schemaName} ===`);

    const exists = await tableExists(token, logicalName);
    if (exists) {
      console.log(`  [skip] table already exists`);
    } else {
      try {
        const body = buildEntityBody(prefix, table);
        console.log(`  [create] ${table.schemaName}`);
        await apiWithRetry(token, "POST", `/EntityDefinitions?SolutionUniqueName=BizzSummit2026`, body);
        console.log(`  [ok] table created`);
        await sleep(500);
      } catch (err) {
        console.log(`  [fail] table ${table.schemaName}: ${err.message.split("\n")[0]}`);
        failures.push({ table: table.schemaName, error: err.message });
        continue;
      }
    }

    for (const attr of table.attributes) {
      const attrLogical = attr.schemaName.toLowerCase();
      const alreadyExists = await attributeExists(token, logicalName, attrLogical);
      if (alreadyExists) {
        console.log(`  [skip] attribute ${attr.schemaName} already exists`);
        continue;
      }

      try {
        if (attr.type === "Lookup") {
          console.log(`  [create lookup] ${attr.schemaName} -> ${attr.targetEntity}`);
          const relBody = buildOneToManyRelationship(prefix, logicalName, attr);
          await apiWithRetry(token, "POST", `/RelationshipDefinitions?SolutionUniqueName=BizzSummit2026`, relBody);
        } else if (attr.type === "PolymorphicLookup") {
          console.log(`  [create polymorphic lookup] ${attr.schemaName} -> [${attr.targets.join(", ")}]`);
          const polyBody = buildPolymorphicLookup(prefix, logicalName, attr);
          await apiWithRetry(token, "POST", `/CreatePolymorphicLookupAttribute`, polyBody);
        } else {
          console.log(`  [create attr] ${attr.schemaName} (${attr.type})`);
          const attrBody = buildAttribute(prefix, attr);
          await apiWithRetry(token, "POST", `/EntityDefinitions(LogicalName='${logicalName}')/Attributes?SolutionUniqueName=BizzSummit2026`, attrBody);
        }
        await sleep(400);
      } catch (err) {
        console.log(`  [fail] ${attr.schemaName}: ${err.message.split("\n")[0]}`);
        failures.push({ table: table.schemaName, attribute: attr.schemaName, error: err.message });
      }
    }
  }

  if (config.standardTableExtensions) {
    for (const ext of config.standardTableExtensions) {
      const entityLogical = ext.entityLogicalName;
      console.log(`\n=== Standard table extension: ${entityLogical} ===`);

      for (const attr of ext.attributes) {
        const attrLogical = attr.schemaName.toLowerCase();
        const alreadyExists = await attributeExists(token, entityLogical, attrLogical);
        if (alreadyExists) {
          console.log(`  [skip] attribute ${attr.schemaName} already exists`);
          continue;
        }
        try {
          console.log(`  [create attr] ${attr.schemaName} (${attr.type})`);
          const attrBody = buildAttribute(prefix, attr);
          await apiWithRetry(
            token,
            "POST",
            `/EntityDefinitions(LogicalName='${entityLogical}')/Attributes?SolutionUniqueName=BizzSummit2026`,
            attrBody
          );
          await sleep(400);
        } catch (err) {
          console.log(`  [fail] ${attr.schemaName}: ${err.message.split("\n")[0]}`);
          failures.push({ table: entityLogical, attribute: attr.schemaName, error: err.message });
        }
      }
    }
  }

  console.log("\n=== Summary ===");
  if (failures.length === 0) {
    console.log("All tables and attributes created successfully.");
  } else {
    console.log(`${failures.length} failures:`);
    for (const f of failures) {
      console.log(`  - ${f.table}${f.attribute ? `.${f.attribute}` : ""}`);
    }
    console.log("\nRe-run the script to retry failed items (existing ones will be skipped).");
  }
  console.log("\nDone.");
}

main().catch((err) => {
  console.error("\nFATAL:", err.message);
  process.exit(1);
});
