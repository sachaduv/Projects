import xml.etree.ElementTree as ET
import re
tree = ET.parse("SyncLUStockAdjustmentMsg.xml")
root = tree.getroot()
print(root)
exp=re.compile(r'{.+}').findall(str(root))[0]
ord=''
ord=root.find('.//'+exp+'BusinessUnitCodeLU')
print('.//'+exp+'BusinessUnitCodeLU')
if(ord not in [None]):
    print(ord.text)
else:
    print("no match")

    
#executable=root.findall('.//{http://ikea.com/ModifyExecutableOrderLine/V2/}ExecutableOrderId')
#print(executable[0].text)
#manual_order_id=root.find('.//{http://ikea.com/RemoveManualOrder/V0/}ManualOrderRequestId')
#print(manual_order_id.text)
#executable = root.findall('.//{http://ikea.com/ExecutableOrder/V2/}ExecutableOrderId')
#print(executable[0].text)


