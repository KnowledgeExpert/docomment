import glob

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree
from dateutil.parser import parse

namespace = {'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


for file in glob.glob("resources/*.docx"):
    print("Checking:", file)
    p = Document(file)
    core_properties = p.core_properties
    print("Created on", core_properties.created)
    cpart = p.part.part_related_by(RT.COMMENTS)
    xmlString = cpart.blob
    root = etree.fromstring(xmlString)
    print("Comments are from:")
    comments = []
    for comment in root.xpath("w:comment", namespaces=namespace):
        date = comment.xpath('@w:date', namespaces=namespace)[0]
        print(date)
        comments.append(parse(date).replace(tzinfo=None))
    oldest = min(comments)
    newest = max(comments)
    print("Oldest:", oldest)
    print("Newest:", newest)
    print("Time passed between first and last comment:", newest - oldest)
    print("Time passed creation and last comment:", newest - core_properties.created)