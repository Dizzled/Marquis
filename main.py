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
## 7. Verification
import xml.etree.ElementTree as ET
import os

class XmlParser:

    def __init__(self, filepath):
        if not os.path.isfile(filepath):
            raise ValueError("File not found")
        self.tree = ET.parse(filepath)
        self.root = self.tree.getroot()
        self.namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        self.findings_dict = {} # Initialize dictionary
        self.i = 1

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
    
    def is_heading3_section(self, p):
        """Returns True if the given paragraph section has been styled as a Heading3"""
        return_val = False
        heading_style_elem = p.find(".//w:pStyle[@w:val='Heading3']", self.namespace)
        if heading_style_elem is not None:
            return_val = True
        return return_val
    
     #
    def is_heading4_section(self, p):
        """Returns 'Heading4Char' if styled as Heading4Char, 'Heading4' if styled as Heading4, and None otherwise"""
        heading4char_elem = p.find(".//w:rStyle[@w:val='Heading4Char']", self.namespace)
        heading4_elem = p.find(".//w:pStyle[@w:val='Heading4']", self.namespace)
        
        if heading4char_elem is not None:
            return 'Heading4Char'
        elif heading4_elem is not None:
            return 'Heading4'
        else:
            return None  # Return None if neither Heading4Char nor Heading4 is found

    def get_section_text(self,p):
        """Returns the joined text of the text elements under the given paragraph tag"""
        return_val = ''
        text_elems = p.findall('.//w:p', self.namespace)
        if text_elems is not None:
            #Neet to set the return_val to only be the items after the first element
            return_val = ''.join([t.text for t in text_elems])
        return return_val

    def get_section4_text(self,p):
        """Returns the joined text of the text elements under the given paragraph tag"""
        return_val = ''
        text_elems = p.findall('.//w:t', self.namespace)
        if text_elems:
            # Set the return_val to only be the items after the first element
            return_val = ''.join([t.text for t in text_elems[1:]]) if len(text_elems) > 1 else ''
        return return_val
    
    def extract_text_after_heading4(self, p, index):
        """Extracts and returns the text from paragraphs following a Heading4 until another Heading4 is encountered."""
        text = []
        paragraphs = self.root.findall('.//w:p', self.namespace)
        for p in paragraphs[index + 1:]:  # Start from the paragraph after the current one
            if self.is_heading4_section(p):  # Stop extraction if another Heading4 is encountered
                break
            else:
                text_elems = p.findall('.//w:t', self.namespace)  # Find all text elements within the paragraph
                paragraph_text = ' '.join([t.text for t in text_elems if t.text is not None])
                text.append(paragraph_text)
        return ' '.join(text)

    
    
    def extract_medium_severity_findings(self):
        paragraphs = self.root.findall('.//w:p', self.namespace)
        medium_severity_section_found = False
        for p in paragraphs:
            if self.is_heading2_section(p) and "Medium Severity Findings" in self.get_section_text(p):
                medium_severity_section_found = True
            elif medium_severity_section_found and self.is_heading3_section(p) and self.get_section_text(p).strip() != '':
                self.findings_dict[f"Title {self.i}"] = self.get_section_text(p)
                self.i = self.i + 1
            elif medium_severity_section_found and self.is_heading2_section(p):
                # another Heading2 found, means we are out of the "Medium Severity Findings" section
                break
        self.extract_low_severity_findings()

    def extract_low_severity_findings(self):
        paragraphs = self.root.findall('.//w:p', self.namespace)
        medium_severity_section_found = False
        for p in paragraphs:
            if self.is_heading2_section(p) and "Low Severity Findings" in self.get_section_text(p):
                medium_severity_section_found = True
            elif medium_severity_section_found and self.is_heading3_section(p) and self.get_section_text(p).strip() != '':
                 self.findings_dict[f"Title {self.i}"] = self.get_section_text(p)
                 self.i = self.i + 1
            elif medium_severity_section_found and self.is_heading2_section(p):
                # another Heading2 found, means we are out of the "Low Severity Findings" section
                break
    
    def extract_high_severity_findings(self):
        paragraphs = self.root.findall('.//w:p', self.namespace)
        high_severity_section_found = False
        current_finding = {}
        attributes_order = ["Severity", "Relevant CWEs", "Vulnerability Details", "Impact", "Recommendation", "Verification"]

        title = None  # Initialize title variable outside of the loop

        for i, p in enumerate(paragraphs):
            if self.is_heading2_section(p) and "High Severity Findings" in self.get_section_text(p):
                high_severity_section_found = True
            elif high_severity_section_found and self.is_heading3_section(p) and self.get_section_text(p).strip() != '':
                if current_finding:  # If there's already a finding being processed, add it to findings_dict
                    title = current_finding.pop("Title", None)
                    if title:  # Check if title is not None before using it as a key
                        self.findings_dict[title] = current_finding  # Using the Title as the key
                    current_finding = {}
                title = self.get_section_text(p)  # Update title variable every time a new Title is found
            
            elif self.is_heading4_section(p): 
               
                heading4_type = self.is_heading4_section(p)
                heading_text = self.get_section_text(p).strip()  # Assuming this method returns the text of the exheading
                for attr in attributes_order:
                    if attr in heading_text:
                        if heading4_type == 'Heading4Char':
                            current_finding[attr] = self.get_section4_text(p)
                        elif heading4_type == 'Heading4':
                            current_finding[attr] = self.extract_text_after_heading4(p, i)


                if title:  # Check if title is not None before using it as a key
                    self.findings_dict[title] = current_finding  # Update the finding with the current attributes and associated text

            elif high_severity_section_found and self.is_heading2_section(p):
                # another Heading2 found, means we are out of the "High Severity Findings" section
                break

        return self.findings_dict


    def print_findings(self):
        for title, finding_details in self.findings_dict.items():
            print(f"Title: {title}")
            for key, value in finding_details.items():
                print(f"\t{key}: {value}")
            print("\n")  
            
                

def main():
    input_file = input("Enter the file name: ")
    try:
        
        parser = XmlParser(input_file)
        parser.remove_hyperlink_tags()
        parser.extract_high_severity_findings()
        parser.print_findings()
        
        #parser.print_body_elements()
    except ValueError as e:
        print(e)


if __name__ == "__main__":
    main()




