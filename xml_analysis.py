from lxml import etree

def analyze_xml_lxml(xml_file):
    try:
        tree = etree.parse(xml_file)
        root = tree.getroot()

        print("Namespaces in the XML:")
        # The nsmap gives a dictionary of prefix-to-URI mappings.
        for prefix, uri in root.nsmap.items():
            if prefix:
                print(f"  Prefix: {prefix} -> URI: {uri}")
            else:
                print(f"  Default Namespace: {uri}")

        print("\nXML Structure Analysis:")
        for element in tree.iter():
            print(f"Tag: {element.tag}")
            if element.attrib:
                print("  Attributes:")
                for attr, value in element.attrib.items():
                    print(f"    {attr} = {value}")
            if element.text and element.text.strip():
                print("  Text: ", element.text.strip())
            print("-" * 40)

    except etree.XMLSyntaxError as e:
        print("Error parsing XML:", e)

# Example usage:
if __name__ == "__main__":
    xml_file = "C:/bcgt/cfx/CFX_24.7/python_project/xml_generator/practice/config_template.xml"
    analyze_xml_lxml(xml_file)