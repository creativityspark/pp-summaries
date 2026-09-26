import { describe, expect, it } from "vitest";
import { prepareFetchXmlForPreview } from "./useDataverseCatalog";

describe("Dataverse record preview", () => {
  it("removes FetchXML limits that conflict with connector paging", () => {
    expect(prepareFetchXmlForPreview('<fetch top="5000"><entity name="account" /></fetch>'))
      .toBe('<fetch><entity name="account" /></fetch>');
  });

  it("keeps view attributes while stripping paging attributes", () => {
    expect(
      prepareFetchXmlForPreview(
        '<fetch version="1.0" output-format="xml-platform" mapping="logical" count="50" page="2" paging-cookie="abc"><entity name="account" /></fetch>',
      ),
    ).toBe('<fetch version="1.0" output-format="xml-platform" mapping="logical"><entity name="account" /></fetch>');
  });

  it("leaves FetchXML without limits untouched", () => {
    const xml = '<fetch distinct="true"><entity name="contact"><attribute name="fullname" /></entity></fetch>';
    expect(prepareFetchXmlForPreview(xml)).toBe(xml);
  });
});
