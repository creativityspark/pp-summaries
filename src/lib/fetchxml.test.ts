import { describe, expect, it } from "vitest";
import { formatFetchXml } from "./fetchxml";

describe("formatFetchXml", () => {
  it("expands single-line saved view FetchXML into indented lines", () => {
    const input =
      '<fetch version="1.0"><entity name="account"><attribute name="name" /><link-entity name="contact" from="contactid" to="primarycontactid"><attribute name="emailaddress1" /></link-entity></entity></fetch>';
    expect(formatFetchXml(input)).toBe(
      [
        '<fetch version="1.0">',
        '  <entity name="account">',
        '    <attribute name="name" />',
        '    <link-entity name="contact" from="contactid" to="primarycontactid">',
        '      <attribute name="emailaddress1" />',
        "    </link-entity>",
        "  </entity>",
        "</fetch>",
      ].join("\n"),
    );
  });

  it("keeps condition values and re-flows already formatted XML", () => {
    const input = '<fetch>\n  <entity name="account">\n<filter><condition attribute="name" operator="eq" value="Contoso Retail" /></filter></entity></fetch>';
    const output = formatFetchXml(input);
    expect(output).toContain('value="Contoso Retail"');
    expect(output.split("\n")).toHaveLength(7);
  });

  it("returns non-XML input unchanged", () => {
    expect(formatFetchXml("not xml")).toBe("not xml");
  });
});
