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
    
    def remove_hyperlink_tags(self):
        for parent in self.root.findall(".//w:r/..", self.namespace):
        # Find all 'w:r' child tags with "HYPERLINK" text under the current parent
            hyperlink_tags = [child for child in parent if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r' and child.find("w:t", self.namespace) is not None and child.find("w:t", self.namespace).text == "HYPERLINK"]
        for tag in hyperlink_tags:
            parent.remove(tag)

    def remove_deleted_text(self):
    # Remove any w:del tags
        for parent in self.root.findall(".//w:del/..", self.namespace):
            del_tags = [child for child in parent if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}del']
        for tag in del_tags:
            parent.remove(tag)

    # Remove any tags with a w:rsidDel attribute
        for parent in self.root.findall(".//*[@w:rsidDel]/..", self.namespace):
            rsid_del_tags = [child for child in parent if "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidDel" in child.attrib]
        for tag in rsid_del_tags:
            parent.remove(tag)
    
    def is_heading2_section(self, p):
        """Returns True if the given paragraph section has been styled as a Heading2"""
        return_val = False
        heading_style_elem = p.find(".//w:pStyle[@w:val='Heading2']", self.namespace)
        if heading_style_elem is not None:
            return_val = True
        return return_val
    
    def get_section_text(self,p):
        """Returns the joined text of the text elements under the given paragraph tag"""
        return_val = ''
        text_elems = p.findall('.//w:t', self.namespace)
        if text_elems is not None:
            return_val = ''.join([t.text for t in text_elems])
        return return_val
    
    def get_paragraph_sections(self):
        """Fetch all paragraph sections."""
        body = self.root.find('.//w:body', self.namespace)
        return body.findall('w:p', self.namespace) if body is not None else []

 



def main():
    input_file = input("Enter the file name: ")
    try:
        
        parser = XmlParser(input_file)
        paragraph_sections = parser.get_paragraph_sections()
        section_labels = [parser.get_section_text(s) if parser.is_heading2_section(s) else '' for s in paragraph_sections]
        print(section_labels)
        parser.remove_hyperlink_tags()
        parser.remove_deleted_text()
        #parser.print_body_elements()
    except ValueError as e:
        print(e)


if __name__ == "__main__":
    main()




