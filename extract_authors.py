#!/usr/bin/env python3

import sys
import xml.etree.ElementTree as ET

def extract_authors(xml_file):
    authors = set()
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        # Namespace handling
        namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        for elem in root.findall('.//*[@w:author]', namespace):
            author = elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author')
            if author:
                authors.add(author)
    except Exception as e:
        print(f"Error processing {xml_file}: {e}", file=sys.stderr)
    return authors

def main():
    if len(sys.argv) != 3:
        print("Usage: extract_authors.py <comments_xml> <document_xml>", file=sys.stderr)
        sys.exit(1)

    comments_file = sys.argv[1]
    document_file = sys.argv[2]

    authors = extract_authors(comments_file) | extract_authors(document_file)
    for author in sorted(authors):
        print(author)

if __name__ == "__main__":
    main()

