import { describe, expect, it } from "vitest";
import {
  buildConfigurationPayload,
  buildContextFetchXml,
  buildPublishSignal,
  buildRecordSelectionFetchXml,
  compileRecipe,
  estimateRecipeTokens,
  type SummaryConfigurationDraft,
} from "./summaryConfiguration";
import * as summaryConfiguration from "./summaryConfiguration";

const draft: SummaryConfigurationDraft = {
  name: "Account operations summary",
  entity: "account",
  sourceFields: ["name", "revenue", "description"],
  outputEntity: "account",
  outputField: "csp_aisummary",
  model: "gpt-4.1-mini",
  mode: 100000000,
  promptId: "98ce2773-0538-4017-bf93-601b01be21eb",
  promptName: "PPS - Generic Summary",
  promptKey: "generic-summary",
  promptContent: "Summarise {{account_context}}.",
  flowName: "csp_SUM_AccountOperations_v3",
  relationships: [{ entity: "activitypointer", windowDays: 30, maxRecords: 12 }],
  inputMappings: { account_context: "compiled.primaryAndRelated" },
  triggerColumns: ["name", "revenue", "description"],
  triggerType: "dataverse.update",
  preserveHistory: false,
  saveMetadata: true,
  queryMode: 100000002,
  maxRecords: 5000,
  fetchXml: "<fetch top=\"5000\"><entity name=\"account\"><attribute name=\"accountid\" /></entity></fetch>",
  relatedFetchXml: "<fetch><entity name=\"account\"><attribute name=\"name\" /></entity></fetch>",
};

describe("summary configuration persistence", () => {
  it("maps the Dataverse table catalog into searchable table metadata", () => {
    const mapTables = (summaryConfiguration as unknown as {
      mapDataverseTables?: (value: unknown) => Array<Record<string, string>>;
    }).mapDataverseTables;

    expect(mapTables).toBeTypeOf("function");
    expect(mapTables?.({ value: [
      { LogicalName: "contact", EntitySetName: "contacts", PrimaryIdAttribute: "contactid", DisplayCollectionName: { UserLocalizedLabel: { Label: "Contacts" } } },
      { LogicalName: "account", EntitySetName: "accounts", PrimaryIdAttribute: "accountid", DisplayCollectionName: { UserLocalizedLabel: { Label: "Accounts" } } },
      { LogicalName: "internal_only" },
    ] })).toEqual([
      { logicalName: "account", entitySetName: "accounts", primaryIdAttribute: "accountid", displayName: "Accounts" },
      { logicalName: "contact", entitySetName: "contacts", primaryIdAttribute: "contactid", displayName: "Contacts" },
    ]);
  });

  it("enriches a selected table with its exact Dataverse identity metadata", () => {
    const mapDefinition = (summaryConfiguration as unknown as {
      mapDataverseTableDefinition?: (value: unknown, fallback: Record<string, string>) => Record<string, string>;
    }).mapDataverseTableDefinition;
    const fallback = { logicalName: "new_person", entitySetName: "new_people", primaryIdAttribute: "new_personid", displayName: "People" };

    expect(mapDefinition).toBeTypeOf("function");
    expect(mapDefinition?.({ data: {
      LogicalName: "new_person",
      EntitySetName: "new_people",
      PrimaryIdAttribute: "new_personkey",
      DisplayCollectionName: { UserLocalizedLabel: { Label: "People" } },
      PrimaryNameAttribute: "new_fullname",
    } }, fallback)).toEqual({
      primaryNameAttribute: "new_fullname",
      logicalName: "new_person",
      entitySetName: "new_people",
      primaryIdAttribute: "new_personkey",
      displayName: "People",
    });
  });

  it("maps Dataverse column metadata and keeps readable labels and capabilities", () => {
    const mapColumns = (summaryConfiguration as unknown as {
      mapDataverseColumns?: (value: unknown) => Array<Record<string, unknown>>;
    }).mapDataverseColumns;

    expect(mapColumns).toBeTypeOf("function");
    expect(mapColumns?.({ Attributes: [
      { LogicalName: "description", AttributeType: "Memo", IsValidForRead: true, IsValidForUpdate: true, DisplayName: { UserLocalizedLabel: { Label: "Description" } } },
      { LogicalName: "accountid", AttributeType: "Uniqueidentifier", IsValidForRead: true, IsValidForUpdate: false, DisplayName: { UserLocalizedLabel: { Label: "Account" } } },
    ] })).toEqual([
      { logicalName: "accountid", displayName: "Account", type: "Uniqueidentifier", readable: true, writable: false },
      { logicalName: "description", displayName: "Description", type: "Memo", readable: true, writable: true },
    ]);
  });

  it("unwraps connector records returned through dynamicProperties", () => {
    const mapRecords = (summaryConfiguration as unknown as {
      mapDataverseRecords?: (value: unknown) => Array<Record<string, unknown>>;
    }).mapDataverseRecords;

    expect(mapRecords).toBeTypeOf("function");
    expect(mapRecords?.({ value: [
      { dynamicProperties: { accountid: "1", name: "Contoso" } },
      { accountid: "2", name: "Northwind" },
    ] })).toEqual([
      { accountid: "1", name: "Contoso" },
      { accountid: "2", name: "Northwind" },
    ]);
  });

  it("maps Dataverse one-to-many relationships into FetchXML join metadata", () => {
    const mapRelationships = (summaryConfiguration as unknown as {
      mapDataverseRelationships?: (value: unknown, sourceEntity: string) => Array<Record<string, unknown>>;
    }).mapDataverseRelationships;

    expect(mapRelationships).toBeTypeOf("function");
    expect(mapRelationships?.({ OneToManyRelationships: [{
      SchemaName: "account_contacts",
      ReferencedEntity: "account",
      ReferencedAttribute: "accountid",
      ReferencingEntity: "contact",
      ReferencingAttribute: "parentcustomerid",
    }] }, "account")).toEqual([{
      schemaName: "account_contacts",
      entity: "contact",
      fromAttribute: "parentcustomerid",
      toAttribute: "accountid",
    }]);
  });

  it("builds a disabled draft row without sending the publish signal", () => {
    const buildDraft = (summaryConfiguration as unknown as {
      buildDraftConfigurationPayload?: (value: SummaryConfigurationDraft) => Record<string, unknown>;
    }).buildDraftConfigurationPayload;

    expect(buildDraft).toBeTypeOf("function");
    expect(buildDraft?.(draft)).toMatchObject({
      csp_name: "Account operations summary",
      csp_enabled: false,
      csp_status: 100000000,
    });
    expect(buildDraft?.(draft)).not.toHaveProperty("statuscode");
  });

  it("reports the fields that would prevent a configuration from publishing", () => {
    const validate = (summaryConfiguration as unknown as {
      validateSummaryConfiguration?: (value: SummaryConfigurationDraft) => string[];
    }).validateSummaryConfiguration;

    expect(validate).toBeTypeOf("function");
    expect(validate?.({
      ...draft,
      name: " ",
      sourceFields: [],
      promptId: "",
      relatedFetchXml: "<fetch><entity name=\"account\" /></fetch>",
      outputField: "",
    })).toEqual([
      "Give this summary a name.",
      "Select at least one source field.",
      "Select an AI Prompt.",
      "The context FetchXML must include {{recordId}}.",
      "Select a destination column.",
    ]);
  });

  it("compiles related-record filters inside the per-record context FetchXML", () => {
    const xml = buildContextFetchXml("account", ["name"], [{
      entity: "activitypointer",
      fetchXml: '<fetch><entity name="activitypointer"><all-attributes /><filter type="and"><condition attribute="statecode" operator="eq" value="0" /></filter></entity></fetch>',
    }]);

    expect(xml).toContain('value="{{recordId}}"');
    expect(xml).toContain('attribute="statecode" operator="eq" value="0"');
    expect(xml).not.toContain('operator="last-x-days"');
  });
  it("compiles the complete user recipe", () => {
    expect(compileRecipe(draft)).toEqual({
      version: "3.0",
      source: {
        entity: "account",
        entitySetName: "accounts",
        primaryIdAttribute: "accountid",
        fields: ["name", "revenue", "description"],
        relationships: [{ entity: "activitypointer", windowDays: 30, maxRecords: 12 }],
        recordSelection: {
          queryMode: 100000002,
          maxRecords: 5000,
          fetchXml: "<fetch top=\"5000\"><entity name=\"account\"><attribute name=\"accountid\" /></entity></fetch>",
        },
        contextFetchXml: "<fetch><entity name=\"account\"><attribute name=\"name\" /></entity></fetch>",
      },
      prompt: {
        id: "98ce2773-0538-4017-bf93-601b01be21eb",
        key: "generic-summary",
        model: "gpt-4.1-mini",
        inputs: { account_context: "compiled.primaryAndRelated" },
      },
      output: { entity: "account", entitySetName: "accounts", field: "csp_aisummary", cache: true, preserveHistory: false },
      trigger: { type: "dataverse.update", columns: ["name", "revenue", "description"] },
    });
  });

  it("persists every searchable field before publishing through statuscode", () => {
    const payload = buildConfigurationPayload(draft);

    expect(payload).toMatchObject({
      csp_name: "Account operations summary",
      csp_targetentity: "account",
      csp_targetentityset: "accounts",
      csp_targetentityidfield: "accountid",
      csp_sourcefields: JSON.stringify(["name", "revenue", "description"]),
      csp_relationships: JSON.stringify([{ entity: "activitypointer", windowDays: 30, maxRecords: 12 }]),
      csp_inputmappings: JSON.stringify({ account_context: "compiled.primaryAndRelated" }),
      csp_outputentity: "account",
      csp_outputfield: "csp_aisummary",
      csp_model: "gpt-4.1-mini",
      csp_mode: 100000000,
      csp_triggercolumns: JSON.stringify(["name", "revenue", "description"]),
      csp_fetchxml: "<fetch top=\"5000\"><entity name=\"account\"><attribute name=\"accountid\" /></entity></fetch>",
      csp_relatedfetchxml: "<fetch><entity name=\"account\"><attribute name=\"name\" /></entity></fetch>",
      csp_querymode: 100000002,
      csp_maxrecords: 5000,
      csp_version: "3.0",
      csp_enabled: true,
      csp_status: 100000001,
      "csp_Prompt@odata.bind": "/csp_aiprompts(98ce2773-0538-4017-bf93-601b01be21eb)",
    });
    expect(payload).not.toHaveProperty("statuscode");
    expect(buildPublishSignal()).toEqual({ statuscode: 787000001 });
    expect(JSON.parse(String(payload.csp_configurationjson))).toEqual(compileRecipe(draft));
  });

  it("persists the exact entity set and primary id returned by Dataverse metadata", () => {
    const payload = buildConfigurationPayload({
      ...draft,
      entity: "new_person",
      entitySetName: "new_people",
      entityIdField: "new_personkey",
    });

    expect(payload).toMatchObject({
      csp_targetentity: "new_person",
      csp_targetentityset: "new_people",
      csp_targetentityidfield: "new_personkey",
    });
  });

  it("keeps the exact destination entity set in the compiled recipe", () => {
    const recipe = compileRecipe({
      ...draft,
      outputEntity: "new_person",
      outputEntitySetName: "new_people",
      outputField: "new_summary",
    });

    expect(recipe.output).toMatchObject({
      entity: "new_person",
      entitySetName: "new_people",
      field: "new_summary",
    });
  });

  it("keeps the selected Dataverse view in the compiled record selection", () => {
    const recipe = compileRecipe({ ...draft, queryMode: 100000000, systemViewId: "view-123", userViewId: undefined });
    expect(recipe.source.recordSelection).toMatchObject({ queryMode: 100000000, viewId: "view-123", viewType: "system" });
  });

  it("builds separate FetchXML queries for record selection and per-record context", () => {
    expect(buildRecordSelectionFetchXml("account", 5000)).toContain('<entity name="account">');
    expect(buildRecordSelectionFetchXml("account", 5000)).toContain('<attribute name="accountid" />');

    const context = buildContextFetchXml("account", ["name", "revenue"], [
      { entity: "activitypointer", windowDays: 30, maxRecords: 12 },
    ]);
    expect(context).toContain('<condition attribute="accountid" operator="eq" value="{{recordId}}" />');
    expect(context).toContain('<attribute name="revenue" />');
    expect(context).toContain('<link-entity name="activitypointer"');
  });

  it("uses the primary id returned by metadata when building FetchXML", () => {
    const build = buildRecordSelectionFetchXml as unknown as (entity: string, maxRecords: number, primaryId: string) => string;
    expect(build("new_person", 25, "new_personkey")).toContain('<attribute name="new_personkey" />');
  });

  it("uses Dataverse relationship attributes when building linked context", () => {
    const context = buildContextFetchXml("account", ["name"], [{
      entity: "contact",
      fromAttribute: "parentcustomerid",
      toAttribute: "accountid",
      fields: ["fullname", "emailaddress1"],
      maxRecords: 10,
    }], "accountid");

    expect(context).toContain('<link-entity name="contact" from="parentcustomerid" to="accountid"');
    expect(context).toContain('<attribute name="fullname" />');
    expect(context).toContain('<attribute name="emailaddress1" />');
  });

  it("estimates tokens from the prompt, fields, and related records", () => {
    const base = estimateRecipeTokens({ promptContent: "", sourceFields: [], relationships: [] });
    const withFields = estimateRecipeTokens({ promptContent: "", sourceFields: ["name", "revenue"], relationships: [] });
    const withRelated = estimateRecipeTokens({
      promptContent: "Summarize {{account_context}} for account managers.",
      sourceFields: ["name", "revenue"],
      relationships: [{ entity: "activitypointer", maxRecords: 12 }],
    });

    expect(base).toBeGreaterThan(0);
    expect(withFields).toBeGreaterThan(base);
    expect(withRelated).toBeGreaterThan(withFields);
    expect(withRelated % 10).toBe(0);
  });
});
