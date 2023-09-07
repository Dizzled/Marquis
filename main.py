#!/usr/bin/env python3
# Import XML File 
# Parse XML File with specific headers
# Required headers
## 1. Title
## 2. Level
## 3. Relevant CWEs
## 4. Vulnerability
## 5. Threat
## 6. Mitigation
## 7.Verification
import xml.etree.ElementTree as ET
import os

class XmlParser:

    def __init__(self, filepath):
        if not os.path.isfile(filepath):
            raise ValueError("File not found")
        self.tree = ET.parse(filepath)
        self.root = self.tree.getroot()
        self.namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    def _print_elements(self, element, indent=0):
        """Recursively print nested XML elements."""
        print(' ' * indent + element.tag.split('}')[-1])  # print the tag without the namespace
        for child in element:
            self._print_elements(child, indent + 2)

    def print_body_elements(self):
        """Find and print elements within the 'w:body' tag."""
        body = self.root.find('.//w:body', self.namespace)
        if body is not None:
            self._print_elements(body)
        else:
            print("w:body tag not found in the XML")
    
    # Remove Hyperlink Tags
    def remove_hyperlink_tags(self):
        # Find all 'w:r' tags with "HYPERLINK" text
        hyperlink_tags = self.root.findall(".//w:r[w:t='HYPERLINK']", self.namespace)
        for tag in hyperlink_tags:
            tag.getparent().remove(tag)


def main():
    input_file = input("Enter the file name: ")
    try:
        parser = XmlParser(input_file)
        parser.print_body_elements()
    except ValueError as e:
        print(e)


if __name__ == "__main__":
    main()




