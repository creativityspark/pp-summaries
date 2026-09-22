import { describe, expect, it } from "vitest";
import {
  buildConfigurationPayload,
  buildContextFetchXml,
  buildPublishSignal,
  buildRecordSelectionFetchXml,
  compileRecipe,
  type SummaryConfigurationDraft,
} from "./summaryConfiguration";

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
      output: { entity: "account", field: "csp_aisummary", cache: true, preserveHistory: false },
      trigger: { type: "dataverse.update", columns: ["name", "revenue", "description"] },
    });
  });

  it("persists every searchable field before publishing through statuscode", () => {
    const payload = buildConfigurationPayload(draft);

    expect(payload).toMatchObject({
      csp_name: "Account operations summary",
      csp_targetentity: "account",
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
});
