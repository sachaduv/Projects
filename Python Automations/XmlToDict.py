from xml.etree import cElementTree as ElementTree

class XmlListConfig(list):
    def __init__(self, aList):
        for element in aList:
            if element:
                if len(element) == 1 or element[0].tag != element[1].tag:
                    self.append(XmlDictConfig(element))
                elif element[0].tag == element[1].tag:
                    self.append(XmlListConfig(element))
            elif element.text:
                text = element.text.strip()
                if text:
                    self.append(text)


class XmlDictConfig(dict):  
    def __init__(self, parent_element):
        if parent_element.items():
            self.update(dict(parent_element.items()))
        for element in parent_element:
            if element:
                if len(element) == 1 or element[0].tag != element[1].tag:
                    aDict = XmlDictConfig(element)
                else:
                    aDict = {element[0].tag: XmlListConfig(element)}
                if element.items():
                    aDict.update(dict(element.items()))
                self.update({element.tag: aDict})
            elif element.items():
                self.update({element.tag: dict(element.items())})
            else:
                self.update({element.tag: element.text})

def convert_xml_to_dict(filename):
    tree = ElementTree.parse(filename+".xml")
    root = tree.getroot()
    return XmlDictConfig(root)


def main():
    xmldict = convert_xml_to_dict('issue_list')
    env = 'issue1'
    conn=[]
    tasks=xmldict[env]['Task_unit']['task']
    for task in tasks:
        conn.append(task['connection_name'])
    print(conn)

if __name__ == "__main__":
    main()