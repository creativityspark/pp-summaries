import { PublicClientApplication } from "@azure/msal-node";

const TENANT_ID = "acb2a8f8-e200-4256-ae26-2104f7620512";
const ENV_URL = "https://spark-tools-dev.crm4.dynamics.com";
const CLIENT_ID = "51f81489-12ee-4a9e-aaae-a2591f45987d";
const SOLUTION_NAME = "BizzSummit2026";

const pca = new PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
  },
});

const token = (await pca.acquireTokenByDeviceCode({
  scopes: [`${ENV_URL}/.default`],
  deviceCodeCallback: (response) => {
    console.log("\nDEVICE CODE AUTHENTICATION");
    console.log(response.message);
    console.log("");
  },
})).accessToken;

async function get(path) {
  const response = await fetch(`${ENV_URL}/api/data/v9.2${path}`, {
    headers: {
      Authorization: `Bearer ${token}`,
      Accept: "application/json",
      "OData-MaxVersion": "4.0",
      "OData-Version": "4.0",
    },
  });
  if (!response.ok) {
    throw new Error(`GET ${path} -> ${response.status}\n${await response.text()}`);
  }
  return response.json();
}

const solutionResult = await get(
  `/solutions?$select=solutionid,uniquename,friendlyname,version,ismanaged&$filter=uniquename eq '${SOLUTION_NAME}'`,
);
const solution = solutionResult.value[0];
if (!solution) throw new Error(`Solution ${SOLUTION_NAME} was not found.`);

const components = await get(
  `/solutioncomponents?$select=componenttype,objectid,rootsolutioncomponentid&$filter=_solutionid_value eq ${solution.solutionid}`,
);

const componentCounts = components.value.reduce((counts, component) => {
  counts[component.componenttype] = (counts[component.componenttype] ?? 0) + 1;
  return counts;
}, {});

const tableComponentIds = new Set(
  components.value
    .filter((component) => component.componenttype === 1)
    .map((component) => component.objectid?.toLowerCase()),
);
const attributeComponentIds = new Set(
  components.value
    .filter((component) => component.componenttype === 2)
    .map((component) => component.objectid?.toLowerCase()),
);

const entities = await get(
  "/EntityDefinitions?$select=MetadataId,LogicalName,SchemaName,EntitySetName,DisplayName,IsCustomEntity,PrimaryIdAttribute,PrimaryNameAttribute",
);
const solutionTables = entities.value
  .filter((entity) => tableComponentIds.has(entity.MetadataId?.toLowerCase()))
  .sort((a, b) => a.LogicalName.localeCompare(b.LogicalName));

const tables = [];
for (const entity of solutionTables) {
  const attributes = await get(
    `/EntityDefinitions(${entity.MetadataId})/Attributes?$select=MetadataId,LogicalName,SchemaName,AttributeType,IsCustomAttribute,RequiredLevel,Description,DisplayName`,
  );
  let recordCount = null;
  try {
    const countResult = await get(`/${entity.EntitySetName}?$count=true&$top=0`);
    recordCount = countResult["@odata.count"] ?? 0;
  } catch {
    recordCount = "unavailable";
  }
  const relationships = await get(
    `/EntityDefinitions(${entity.MetadataId})/ManyToOneRelationships?$select=MetadataId,SchemaName,ReferencedEntity,ReferencingEntity,ReferencingAttribute,IsCustomRelationship`,
  );
  tables.push({
    logicalName: entity.LogicalName,
    schemaName: entity.SchemaName,
    entitySetName: entity.EntitySetName,
    isCustomEntity: entity.IsCustomEntity,
    primaryId: entity.PrimaryIdAttribute,
    primaryName: entity.PrimaryNameAttribute,
    recordCount,
    attributes: attributes.value
      .filter((attribute) =>
        entity.IsCustomEntity
          ? attribute.IsCustomAttribute
          : attributeComponentIds.has(attribute.MetadataId?.toLowerCase()),
      )
      .map((attribute) => ({
        logicalName: attribute.LogicalName,
        schemaName: attribute.SchemaName,
        type: attribute.AttributeType,
        isCustom: attribute.IsCustomAttribute,
        required: attribute.RequiredLevel?.Value,
      }))
      .sort((a, b) => a.logicalName.localeCompare(b.logicalName)),
    relationships: relationships.value
      .filter((relationship) => relationship.IsCustomRelationship)
      .map((relationship) => ({
        schemaName: relationship.SchemaName,
        referencedEntity: relationship.ReferencedEntity,
        referencingEntity: relationship.ReferencingEntity,
        referencingAttribute: relationship.ReferencingAttribute,
      }))
      .sort((a, b) => a.schemaName.localeCompare(b.schemaName)),
  });
}

console.log(JSON.stringify({
  environment: ENV_URL,
  solution,
  componentCounts,
  tables,
}, null, 2));
