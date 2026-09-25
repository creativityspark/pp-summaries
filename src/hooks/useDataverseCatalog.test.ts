import { describe, expect, it } from "vitest";
import { limitFetchXmlForPreview } from "./useDataverseCatalog";

describe("Dataverse record preview", () => {
  it("places the preview limit in FetchXML instead of requiring an incompatible $top parameter", () => {
    expect(limitFetchXmlForPreview('<fetch version="1.0"><entity name="account" /></fetch>', 10))
      .toBe('<fetch version="1.0" top="10"><entity name="account" /></fetch>');
  });

  it("replaces an existing FetchXML top value", () => {
    expect(limitFetchXmlForPreview('<fetch top="5000"><entity name="account" /></fetch>', 10))
      .toBe('<fetch top="10"><entity name="account" /></fetch>');
  });
});
