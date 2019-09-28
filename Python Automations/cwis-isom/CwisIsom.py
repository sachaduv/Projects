import cx_Oracle
import xlsxwriter
import xlrd
import os
from datetime import datetime
from datetime import timedelta
from tkinter import * 
import re
import time
import xml.etree.ElementTree as ET
from win32com.client import Dispatch


screen = Tk()
screen.title("CWIS-ISOM")
screen.resizable(0, 0)
screen.configure(background='light yellow') 
# scr_wdh,scr_ht=screen.winfo_screenwidth(),screen.winfo_screenheight()
# screen.geometry("%dx%d+0+0" % (scr_wdh/2+50, scr_ht/2))
# print(scr_wdh,scr_ht)
#screen.geometry("650x150+%d+%d" %( ( (screen.winfo_screenwidth() / 2.) - (350 / 2.) ), ( (screen.winfo_screenheight() / 2.) - (150 / 2.) ) ) )
var1=0
e1=StringVar()
e2=StringVar()
text1=Entry(screen,bg="light green",width=40,state=DISABLED,textvariable=e1)
text2=Entry(screen,bg="light blue",width=40,state=DISABLED,textvariable=e2)
process=''
input_btn=''
req_type=''
env=''
day=''
iip_obj=''
wis_data=''
wis_logic=''
status=0
wb_r=''
wb=''
display_area=Text(screen,height=10,state=DISABLED)
display_area.grid(row=3,column=0,columnspan=3)

def display(text):
    global display_area
    display_area.config(state=NORMAL)
    display_area.delete("1.0",END)
    display_area.tag_config("tag-center",justify=CENTER)
    display_area.insert("1.0",text,'tag-center')
    display_area.config(state=DISABLED)
    screen.update_idletasks()

def clearDisplay():
    global display_area
    display_area.config(state=NORMAL)
    display_area.delete("1.0",END)
    display_area.config(state=DISABLED)
    screen.update_idletasks()
def closeInput():
    pass
def closeOutput():
    try:
        os.system('TASKKILL /F /IM excel.exe')
    except:
        pass

def readInput():
    global wb_r
    wb_r = xlrd.open_workbook("input.xlsx")

def writeOutput():
    global wb
    wb = xlsxwriter.Workbook("output.xlsx")

def getConnections():
    global env,iip_obj,wis_data,wis_logic
    iip_obj = cx_Oracle.connect("IIP_OBJ","O0144670",env)
    wis_data = cx_Oracle.connect("WIS_DATA","h22U9076",env)
    wis_logic  = cx_Oracle.connect("WIS_LOGIC","p4766U4U",env)

def closeConnections():
    iip_obj.close()
    wis_data.close()
    wis_logic.close()

def deleteContent(fName):
    with open(fName, "w"):
        pass

def checkLogs(ws_log,row,msg_name,tag_name,data_in,data_out):
    global iip_obj
    xml_message_log=iip_obj.cursor()
    xml_message_log.execute("SELECT xml.TRANS_ID,xml.TRANS_STATUS,xml.XML_MESSAGE.GETCLOBVAL(),xml.ERROR_MESSAGE,xml.INS_DATE FROM XML_MESSAGE_LOG xml WHERE xml.MESSAGE_TYPE IN ('?xml','"+msg_name+"') AND TRUNC(xml.INS_DATE) BETWEEN (SYSDATE-4) AND SYSDATE ORDER BY xml.INS_DATE DESC")
    xml_message_log_data = xml_message_log.fetchall()
    log_status=False
    if(xml_message_log.rowcount>0):
        for log in xml_message_log_data:
            with open("xml_msg.xml","w") as x:
                x.write(log[2].read())
            trans_id = log[0]
            trans_status=log[1]
            err_msg = log[3]
            ins_date = log[4].strftime("%d-%b-%y %I.%M.%S.%f %p")
            log_tree = ET.parse("xml_msg.xml")
            log_root=log_tree.getroot()
            log_root_exp=re.compile(r'{.+}').findall(str(log_root))[0]
            log_ord=log_root.find('.//'+log_root_exp+tag_name)
            if(log_ord not in [None]):
                if(log_ord.text in data_in):
                    ws_log.write(0,0,"Status")
                    ws_log.write(0,1,tag_name)
                    ws_log.write(0,2,"TRANS_ID")
                    ws_log.write(0,3,"TRANS_STATUS")
                    ws_log.write(0,4,"ERROR_MESSAGE")
                    ws_log.write(0,5,"INS_DATE")
                    ws_log.write(row,0,data_out)
                    ws_log.write(row,1,log_ord.text)
                    ws_log.write(row,2,trans_id)
                    ws_log.write(row,3,trans_status)
                    ws_log.write(row,4,err_msg)
                    ws_log.write(row,5,ins_date)
                    log_status=True
                    break
            else:
                ws_log.write(row,0,"NameSpace is incorrect for "+msg_name)

        xml_message_log.close()

        if(log_status==False):
            ws_log.write(row,0,"No entry found for "+tag_name+" "+data_in[0])
            return False
        else:
            return True
    else:
        ws_log.write(row,0,"No Transactions in Logs for "+msg_name)
        return False
    

def checkQueues(ws_que,row,que_name,tag_name,data_in,data_out):
    global iip_obj
    q_ifc = iip_obj.cursor()
    q_ifc.execute("SELECT q.USER_DATA.GETCLOBVAL(),q.ENQ_TIME FROM "+que_name+" q WHERE TRUNC(q.ENQ_TIME) BETWEEN (SYSDATE-4) AND SYSDATE ORDER BY q.ENQ_TIME DESC")
    q_ifc_data=q_ifc.fetchall()
    log_status=False
    if(q_ifc.rowcount>0):
        for que in q_ifc_data:
            with open("xml_msg.xml","w") as xml:
                xml.write(que[0].read())
            enq_time=que[1].strftime("%d-%b-%y %I.%M.%S.%f %p")
            que_tree = ET.parse("xml_msg.xml")
            que_root=que_tree.getroot()
            que_root_exp=re.compile(r'{.+}').findall(str(que_root))[0]
            que_ord=que_root.find('.//'+que_root_exp+tag_name)
            if(que_ord not in [None]):
                if(que_ord.text in data_in):
                    ws_que.write(0,0,"Status")
                    ws_que.write(0,1,tag_name)
                    ws_que.write(0,2,"ENQ_TIME")
                    ws_que.write(row,0,data_out)
                    ws_que.write(row,1,que_ord.text)
                    ws_que.write(row,2,enq_time)
                    log_status=True
                    break
            else:
                ws_que.write(row,0,"NameSpace is incorrect for "+que_name)

        q_ifc.close()

        if(log_status==False):
            ws_que.write(row,0,"No entry found for "+tag_name+" "+data_in[0])
            return False
        else:
            return True
    

class Orderprocess:

    def __init__(self,wb,wb_r):
        self.wb=wb
        self.wb_r=wb_r

    def processOrders(self):
        global var1,e1,e2
        #Reading orders from input.xlsx
        ord=''
        orders = self.wb_r.sheet_by_name("Order Processing")
        sales_orders = []
        if(var1==0):
            for i in range(1,orders.nrows):
                sales_orders.append(orders.cell_value(i,0))
        else:
            sales_orders = e1.get().split(",")

        for index in range(len(sales_orders)):
                ord+="'"+sales_orders[index]+"'"
                if(index<len(sales_orders)-1):
                    ord+=","

        ws_iip = self.wb.add_worksheet("I_ORDERS")
        carry_on = True
        
        #Creating Cursor variables
        i_orders = iip_obj.cursor()
        i_order_status = iip_obj.cursor()
        provide_order_lu = iip_obj.cursor()
        order_lines = wis_data.cursor()
        ss_results = wis_data.cursor()
        ss_candidates = wis_data.cursor()
        stock_act_lu_vw = wis_data.cursor()
        stock_alloc_failure_reason_vw = wis_data.cursor()
        results = wis_data.cursor()
        order_status = wis_data.cursor()   
        process_orders = wis_logic.cursor()
         
        display("Checking Order Processing Detais")
        i_orders.execute("SELECT ORD_ID_TK,BU_CODE_LU,WRK_ORD_REF,ORG_ORD_REF,ORD_ID_REF_SALES,TRANS_STATUS,INS_DATE,PAYMENT_OPTION FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+")")
        i_orders_data = i_orders.fetchall()
        if(i_orders.rowcount==0):
            display("No orders found :( checking logs!..")
            ws_log=self.wb.add_worksheet("XML_MESSAGE_LOG")
            carry_on=checkLogs(ws_log,1,'SyncReleaseOrderMsg','OriginalSalesOrderReference',sales_orders,'Order is stuck in XML_MESSAGE_LOG')
            #carry_on=checkLogs(ws_log,1,'?xml','WorkOrderReference',['375549'],'Order is stuck in XML_MESSAGE_LOG') 
            if(not carry_on):
                display("checking in_queue")
                ws_que=self.wb.add_worksheet("Queue")
                carry_on=checkQueues(ws_que,1,'Q_IFC_IN_001','OriginalSalesOrderReference',sales_orders,'Orders are stuck in queue')
            if(not carry_on):
                display("No Orders found")
                ws_iip.write(0,0,"No Orders found")
                carry_on = False
        if(carry_on):
            display("checking i_orders")
            ord_chg = False
            carry_on = False
            track_cdc = []
            for row in i_orders_data:
                cdc = row[1]
                if(row[5]==0):
                    if(cdc not in track_cdc):
                        try:
                            display("Trans_Status is 0 running orders_api.process_orders in wis_logic for cdc "+cdc)
                            process_orders.callproc("orders_api.process_orders",[cdc,'CDC'])
                            wis_logic.commit()
                        except Exception as e:
                            with open("log.txt", 'a') as out:
                                out.write(str(e))
                            return 0
                        try:
                            display("Trans_Status is 0 running ifc_orders_api.provide_order_lu in iip_obj for cdc "+cdc)
                            provide_order_lu.callproc("ifc_orders_api.provide_order_lu",[cdc,"CDC"])
                            iip_obj.commit()
                        except Exception as e:
                            with open("log.txt", 'a') as out:
                                out.write(str(e))
                            return 0
                    ord_chg = True
                    track_cdc.append(cdc)
            ord_r=0
            if(ord_chg):
                i_orders.execute("SELECT ORD_ID_TK,BU_CODE_LU,WRK_ORD_REF,ORG_ORD_REF,ORD_ID_REF_SALES,TRANS_STATUS,INS_DATE,PAYMENT_OPTION FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+")")
                i_orders_data=i_orders.fetchall()
            ws_iip.write(0,0,"ORD_ID_TK")
            ws_iip.write(0,1,"CDC")
            ws_iip.write(0,2,"WRK_ORD_REF")
            ws_iip.write(0,3,"ORG_ORD_REF")
            ws_iip.write(0,4,"ORD_ID_REF_SALES")
            ws_iip.write(0,5,"TRANS_STATUS")
            ws_iip.write(0,6,"INS_DATE")  
            ws_iip.write(0,7,"PAYMENT_OPTION") 
            ws_iip.write(0,8,"ORD_STATUS")            
            for order in i_orders_data:
                ord_r=ord_r+1
                ord_id_tk = str(order[0])
                cdc = order[1]
                wrk_ord_ref = order[2]
                org_ord_ref=order[3]
                ord_id_ref_sales = order[4]
                trans_status = str(order[5])        
                ins_date = order[6].strftime("%d-%b-%y %I.%M.%S.%f %p")
                ins_date=ins_date+" +00:00"
                payment_option = order[7]
                i_ord_status = ''
            
                if(order[5]==7):
                    carry_on = True
                if(not trans_status == '7'):
                    if(trans_status == '5'):
                        trans_status = 'trans_status is 5!..Error please check the logs'
                    elif(trans_status == '0'):
                        trans_status = "trans_status is 0!..Error in processing the order"
                    else:
                        trans_status = "trans_status is "+trans_status+" invalid!.. please report developer"
                try:
                    i_order_status.execute("SELECT ORD_STATUS FROM I_ORDER_STATUS WHERE ORG_ORD_REF= '"+str(org_ord_ref)+"' ORDER BY INS_DATE DESC FETCH FIRST 1 ROWS ONLY")
                    i_order_status_data= i_order_status.fetchone()
                    if(i_order_status.rowcount==0):
                        i_ord_status = 'NULL'
                    else:
                        i_ord_status = i_order_status_data[0]
                except:
                    i_ord_status = 'I_ORDER_STATUS table or view does not exist'
                ws_iip.write(ord_r,0,ord_id_tk)
                ws_iip.write(ord_r,1,cdc)
                ws_iip.write(ord_r,2,wrk_ord_ref)
                ws_iip.write(ord_r,3,org_ord_ref)
                ws_iip.write(ord_r,4,ord_id_ref_sales)
                ws_iip.write(ord_r,5,trans_status)
                ws_iip.write(ord_r,6,ins_date)
                ws_iip.write(ord_r,7,payment_option)
                ws_iip.write(ord_r,8,i_ord_status)

        process_orders.close()
        if(carry_on):
            display("checking order lines")
            ws_data = self.wb.add_worksheet("ORDER_LINES")
            ordl_r = 0

            order_lines.execute("SELECT ORD_ID,ITEM_NO,ITEM_QTY,ORDL_STATUS,INS_DATE,ORD_ID_IN_TK,LU_TK,SS_RUN_TK,CANCEL_SRC,BU_CODE_SUP FROM ORDER_LINES WHERE ORD_ID_IN_TK IN (SELECT ORD_ID_TK FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+"))")
            order_lines_data = order_lines.fetchall()
            if(order_lines.rowcount==0):
                ws_data.write(0,0,"No records in Order Lines")
                carry_on = False
            if(carry_on):
                carry_on = False
                track_cdc=[]
                ordl_chg = False
                for row in order_lines_data:
                    if(row[3]==20):
                        cdc=''
                        i_orders.execute("SELECT BU_CODE_LU FROM I_ORDERS WHERE ORD_ID_TK="+str(row[5]))
                        i_orders_data = i_orders.fetchall()
                        for bu_code_lu in i_orders_data:
                            cdc = bu_code_lu[0]
                        if(cdc not in track_cdc):
                            try:
                                display("Order Line Status is 20 running ifc_orders_api.provide_order_lu in iip_obj for cdc "+cdc)
                                provide_order_lu.callproc("ifc_orders_api.provide_order_lu",[cdc,"CDC"])
                                iip_obj.commit()
                            except Exception as e:
                                with open("log.txt", 'a') as out:
                                    out.write(str(e))
                                return 0
                        ordl_chg = True
                        track_cdc.append(cdc)
                ws_data.write(0,0,"ORD_ID")
                ws_data.write(0,1,"ITEM_NO")
                ws_data.write(0,2,"ITEM_QTY")
                ws_data.write(0,3,"ORDL_STATUS")
                ws_data.write(0,4,"INS_DATE")
                ws_data.write(0,5,"ORD_ID_IN_TK")
                ws_data.write(0,6,"LU_TK")
                ws_data.write(0,7,"SS_RUN_TK")
                ws_data.write(0,8,"ITEM_QTY_AVAIL")
                ws_data.write(0,9,"LAST_CHANGED_BY")
                ws_data.write(0,10,"ELIMINATION_REASON")
                ws_data.write(0,11,"REASON_CODE")
                ws_data.write(0,12,"ORD_STATUS")
                ws_data.write(0,13,"CANCEL_SRC")
                ws_data.write(0,14,"BU_CODE_SUP")
                if(ordl_chg):
                    order_lines.execute("SELECT ORD_ID,ITEM_NO,ITEM_QTY,ORDL_STATUS,INS_DATE,ORD_ID_IN_TK,LU_TK,SS_RUN_TK,CANCEL_SRC,BU_CODE_SUP FROM ORDER_LINES WHERE ORD_ID_IN_TK IN (SELECT ORD_ID_TK FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+"))")
                    order_lines_data = order_lines.fetchall()
                for ordl in order_lines_data:
                    ordl_r+=1                       
                    ord_id = ordl[0]
                    item_no = str(ordl[1])
                    item_qty = ordl[2]
                    ins_date = ordl[4].strftime("%d-%b-%y %I.%M.%S.%f %p")
                    ins_date=ins_date+" +00:00"
                    ord_id_in_tk = str(ordl[5])
                    lu_tk = str(ordl[6])
                    ss_run_tk=str(ordl[7])
                    last_changed_by=''
                    elimination_reason=''
                    reason_code=''
                    ord_status = ''
                    cancel_src = str(ordl[8])
                    bu_code_sup = str(ordl[9])
                    if(ss_run_tk not in ['None']):
                        ss_results.execute("SELECT LAST_CHANGED_BY FROM SS_RESULTS WHERE SS_RUN_TK= "+str(ss_run_tk)+" FETCH FIRST 1 ROWS ONLY")
                        ss_results_data=ss_results.fetchone()
                        if(ss_results.rowcount==0):
                            last_changed_by='NULL'
                        else:
                            last_changed_by=ss_results_data[0]
                        ss_candidates.execute("SELECT ELIMINATION_REASON FROM SS_CANDIDATES WHERE SS_RUN_TK= "+str(ss_run_tk)+" FETCH FIRST 1 ROWS ONLY")
                        ss_candidates_data=ss_candidates.fetchone()
                        if(ss_candidates.rowcount==0):
                            elimination_reason='NULL'
                        else:
                            elimination_reason=ss_candidates_data[0]
                        try:
                            stock_alloc_failure_reason_vw.execute("SELECT REASON_CODE FROM STOCK_ALLOC_FAILURE_REASON_VW WHERE SS_RUN_TK= "+str(ss_run_tk)+" FETCH FIRST 1 ROWS ONLY")
                            stock_alloc_failure_reason_vw_data = stock_alloc_failure_reason_vw.fetchone()
                            if(stock_alloc_failure_reason_vw.rowcount == 0):
                                reason_code = 'NULL'
                            else:
                                reason_code = stock_alloc_failure_reason_vw_data[0]
                        except:
                            reason_code = 'STOCK_ALLOC_FAILURE_REASON_VW table or view does not exist'
                    stock_act_lu_vw.execute("SELECT sum(item_qty-item_qty_stop_tot-item_qty_damaged-item_qty_alloc-item_qty_resv) AS item_qty_avail FROM STOCK_ACT_LU_VW WHERE ITEM_NO = "+item_no+"AND LU_TK ="+lu_tk)
                    stock_act_lu_vw_data = stock_act_lu_vw.fetchone()
                    item_qty_avail=stock_act_lu_vw_data[0]
                    if(item_qty_avail is None):
                        item_qty_avail='None'

                    if(ordl[3] in (30,40)):
                        ordl_status = 'Processed to ASTRO'                       
                        carry_on = True
                    elif(ordl[3]==90):
                        if(item_qty_avail == 'None'):
                            ordl_status = 'Cancelled at CWIS, due to no stock'
                        elif((item_qty_avail != 'None') and (item_qty_avail<item_qty)):
                            ordl_status = 'Cancelled at CWIS, due to no stock'                           
                        else:
                            ordl_status = 'Cancelled at CWIS'
                        carry_on = True
                    elif(row[3]==80):
                        ordl_status = 'Loaded to ASTRO'
                        carry_on = True
                    elif(row[3]==20):
                        ordl_status = 'ordl_status is 20!..Error on job processing'
                    else:
                        ordl_status = 'Error occured!...'
                    try:
                        order_status.execute("SELECT ORD_STATUS FROM ORDER_STATUS WHERE ORD_ID= "+str(ord_id)+" ORDER BY INS_DATE DESC FETCH FIRST 1 ROWS ONLY")
                        order_status_data = order_status.fetchone()
                        if(order_status.rowcount==0):
                            ord_status = 'NULL'
                        else:
                            ord_status = order_status_data[0]
                    except:
                        ord_status = "ORDER_STATUS table or view does not exist"
                    ws_data.write(ordl_r,0,ord_id)
                    ws_data.write(ordl_r,1,item_no)
                    ws_data.write(ordl_r,2,item_qty)
                    ws_data.write(ordl_r,3,ordl_status)
                    ws_data.write(ordl_r,4,ins_date)
                    ws_data.write(ordl_r,5,ord_id_in_tk)
                    ws_data.write(ordl_r,6,lu_tk)
                    ws_data.write(ordl_r,7,ss_run_tk)
                    ws_data.write(ordl_r,8,item_qty_avail)
                    ws_data.write(ordl_r,9,last_changed_by)
                    ws_data.write(ordl_r,10,elimination_reason)
                    ws_data.write(ordl_r,11,reason_code)
                    ws_data.write(ord_r,12,ord_status)
                    ws_data.write(ordl_r,13,cancel_src)
                    ws_data.write(ordl_r,14,bu_code_sup)
                
        i_orders.close()
        order_lines.close()
        provide_order_lu.close()
        stock_act_lu_vw.close()
        stock_alloc_failure_reason_vw.close()
        i_order_status.close()
        order_status.close()
        if(carry_on):
            display("Fetching Order Results")
            ws_results = self.wb.add_worksheet("Order Processing Results")
            res_r = 0

            ws_results.write(0,0,"Sales Order")
            ws_results.write(0,1,"CDC")
            ws_results.write(0,2,"Work Order")
            ws_results.write(0,3,"Item number")
            ws_results.write(0,4,"Status at CWIS")
            ws_results.write(0,5,"Executable Order id")
                       
            results.execute("SELECT i_ord.ORD_ID_REF_SALES,i_ord.BU_CODE_LU,i_ord.WRK_ORD_REF,ord_lns.ITEM_NO,ord_lns.ORDL_STATUS,ord_lns.ORD_ID FROM I_ORDERS i_ord,ORDER_LINES ord_lns WHERE i_ord.ORD_ID_TK = ord_lns.ORD_ID_IN_TK AND i_ord.ORD_ID_REF_SALES IN ("+ord+")")
            for row in results:
                res_r = res_r+1
                rs_ord_id_ref_sales = row[0]
                rs_cdc = row[1]
                rs_wrk_ord_ref = row[2]
                rs_item_no = row[3]
                rs_ord_id = str(row[5])

                if(row[4] in (30,40)):
                    rs_ordl_status = 'Processed to ASTRO'
                elif(row[4]==90):
                    rs_ordl_status = 'Cancelled at CWIS, due to no stock'
                elif(row[4]==80):
                    rs_ordl_status = 'Loaded to ASTRO'
                else:
                    rs_ordl_status = "Error Occured!..."

                ws_results.write(res_r,0,rs_ord_id_ref_sales)
                ws_results.write(res_r,1,rs_cdc)
                ws_results.write(res_r,2,rs_wrk_ord_ref)
                ws_results.write(res_r,3,rs_item_no)
                ws_results.write(res_r,4,rs_ordl_status)
                ws_results.write(res_r,5,rs_ord_id)
        results.close()
        return 1

class InventoryDetails:
    
    def __init__(self,wb,wb_r):
        self.wb = wb
        self.wb_r = wb_r

    def processInventory(self):
        global var1,e1,e2
        stock_r = self.wb_r.sheet_by_name("Inventory Details")
        cdc = ''
        items = []
        articles = ''
        if(var1==0):
            cdc = stock_r.cell_value(1,0)
            for i in range(1,stock_r.nrows):
                items.append(stock_r.cell_value(i,1))
        else:
            cdc=e1.get()
            items=e2.get().split(",")
        for i in range(len(items)):
            articles+="'"+items[i]+"'"
            if(i<len(items)-1):
                articles+=","

        provide_stock_cos = wis_logic.cursor()
        o_stock_cos = iip_obj.cursor()
        
        ws_inventory = self.wb.add_worksheet("Inventory Details_"+cdc+"_"+env)
        if(articles == "''"):
            o_stock_cos.execute("SELECT * FROM O_STOCK_COS WHERE BU_CODE_LU = '"+cdc+"' AND TRUNC(INS_DATE) LIKE TRUNC(sysdate) ORDER BY ITEM_QTY DESC fetch first 50 rows only") 
            o_stock_cos_data = o_stock_cos.fetchall()
        else:
            o_stock_cos.execute("SELECT * FROM O_STOCK_COS WHERE BU_CODE_LU = '"+cdc+"' AND TRUNC(INS_DATE) LIKE TRUNC(sysdate) AND ITEM_NO IN ("+articles+") ORDER BY ITEM_QTY DESC") 
            o_stock_cos_data = o_stock_cos.fetchall()

        if(o_stock_cos.rowcount==0):
            try:
                display("running IFC_STOCK_API.PROVIDE_STOCK_COS job for cdc "+cdc)
                provide_stock_cos.callproc('IFC_STOCK_API.PROVIDE_STOCK_COS',[cdc,'CDC'])
                wis_logic.commit()
            except Exception as e:
                with open("mail_body.txt", 'a') as out:         
                    out.write(str(e))
                return 0
            if(articles == "''"):
                o_stock_cos.execute("SELECT * FROM O_STOCK_COS WHERE BU_CODE_LU = '"+cdc+"' AND TRUNC(INS_DATE) LIKE TRUNC(sysdate) ORDER BY ITEM_QTY DESC fetch first 50 rows only") 
                o_stock_cos_data = o_stock_cos.fetchall()
            else:
                o_stock_cos.execute("SELECT * FROM O_STOCK_COS WHERE BU_CODE_LU = '"+cdc+"' AND TRUNC(INS_DATE) LIKE TRUNC(sysdate) AND ITEM_NO IN ("+articles+") ORDER BY ITEM_QTY DESC") 
                o_stock_cos_data = o_stock_cos.fetchall()
        display("Fetching Stock Details...")
        if(o_stock_cos.rowcount==0 and day=='Today'):
            ws_inventory.write(0,0,"No Inventory on today's date")
            return 1
        elif(o_stock_cos.rowcount>0):
            pass
        else:
            if(articles == "''"):
                o_stock_cos.execute("SELECT * FROM O_STOCK_COS WHERE BU_CODE_LU = '"+cdc+"'  ORDER BY INS_DATE DESC,ITEM_QTY DESC fetch first 50 rows only") 
                o_stock_cos_data = o_stock_cos.fetchall()
            else:
                o_stock_cos.execute("SELECT * FROM O_STOCK_COS WHERE BU_CODE_LU = '"+cdc+"'  AND TRUNC(INS_DATE) LIKE TRUNC(sysdate-1) AND ITEM_NO IN ("+articles+") ORDER BY INS_DATE DESC,ITEM_QTY DESC") 
                o_stock_cos_data = o_stock_cos.fetchall()

        inv_r = 0
        ws_inventory.write(0,0,"BU_CODE_LU")
        ws_inventory.write(0,1,"BU_TYPE_LU")
        ws_inventory.write(0,2,"ITEM_NO")
        ws_inventory.write(0,3,"ITEM_TYPE")
        ws_inventory.write(0,4,"BU_CODE_SUP")
        ws_inventory.write(0,5,"BU_TYPE_SUP")
        ws_inventory.write(0,6,"TRANS_DATE")
        ws_inventory.write(0,7,"ITEM_QTY")
        ws_inventory.write(0,8,"UOM_CODE") 
            
        for row in o_stock_cos_data:
            inv_r +=1
            bu_code_lu = row[1]
            bu_type_lu = row[2]
            item_no = row[3]
            item_type = row[4]
            bu_code_sup = row[5]
            bu_type_sup = row[6]
            trans_date = row[7].strftime("%d-%b-%Y")
            item_qty = row[8]
            uom_code = row[9] 
            ws_inventory.write(inv_r,0,bu_code_lu)
            ws_inventory.write(inv_r,1,bu_type_lu)
            ws_inventory.write(inv_r,2,item_no)
            ws_inventory.write(inv_r,3,item_type)
            ws_inventory.write(inv_r,4,bu_code_sup)
            ws_inventory.write(inv_r,5,bu_type_sup)
            ws_inventory.write(inv_r,6,trans_date)
            ws_inventory.write(inv_r,7,item_qty)
            ws_inventory.write(inv_r,8,uom_code)

        provide_stock_cos.close()
        o_stock_cos.close()
        return 1

class PublishStockISOM:
    def __init__(self,wb,wb_r,onhand):
        self.wb=wb
        self.wb_r=wb_r 
        self.onhand = onhand

    def processStockToISOM(self):
        global day,var1,e1,e2
        bu_code = self.wb_r.sheet_by_name("Publish Stock to ISOM")
        publish_cdc = []
        if(var1==0):
            for i in range(1,bu_code.nrows):
                publish_cdc.append(bu_code.cell_value(i,0))
        else:
            publish_cdc=e1.get().split(",")

        provide_stock_isom = wis_logic.cursor()
        i_stock_reps=iip_obj.cursor()
        transaction_ctrl=iip_obj.cursor()
        o_stock_lu_cust_prom_th = iip_obj.cursor()
        o_stock_lu_cust_prom_tl = iip_obj.cursor()
        
        for cdc in publish_cdc:
            carry_on = False
            if(not self.onhand):
                display("checking i_stock_reps")
                stk_reps = self.wb.add_worksheet("I_STOCK_REPS_"+cdc)
                if(day=='Today'):
                    i_stock_reps.execute("SELECT BU_CODE_LU,TRANS_DATE,TRANS_STATUS,ERR_REASON,INS_DATE FROM I_STOCK_REPS WHERE BU_CODE_LU='"+cdc+"' AND TRUNC(INS_DATE) LIKE TRUNC(sysdate) ORDER BY INS_DATE DESC")
                    i_stock_reps_data = i_stock_reps.fetchall()
                else:
                    i_stock_reps.execute("SELECT BU_CODE_LU,TRANS_DATE,TRANS_STATUS,ERR_REASON,INS_DATE FROM I_STOCK_REPS WHERE BU_CODE_LU='"+cdc+"'  ORDER BY INS_DATE DESC")
                    i_stock_reps_data=i_stock_reps.fetchall()
                if(i_stock_reps.rowcount>0):
                    for reps in i_stock_reps_data:
                        if(reps[2] in (1,7)):             
                            bu_code_lu=reps[0]
                            trans_date=reps[1].strftime("%d-%b-%Y")
                            trans_status = reps[2]
                            err_reason=reps[3]
                            ins_date = reps[4].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00"
                            stk_reps.write(0,0,"BU_CODE_LU")
                            stk_reps.write(0,1,"TRANS_DATE")
                            stk_reps.write(0,2,"TRANS_STATUS")
                            stk_reps.write(0,3,"ERR_REASON")
                            stk_reps.write(0,4,"INS_DATE")
                            stk_reps.write(1,0,bu_code_lu)
                            stk_reps.write(1,1,trans_date)
                            stk_reps.write(1,2,trans_status)
                            stk_reps.write(1,3,err_reason)
                            stk_reps.write(1,4,ins_date) 
                            carry_on=True
                            break
                        else:
                            stk_reps.write(0,0,"Error!..TRANS_STATUS is "+str(reps[2])+" for cdc "+str(cdc))
                            carry_on=False
                            break 
                else:
                    stk_reps.write(0,0,"No Stock Report for cdc "+cdc)
            if(self.onhand):
                carry_on=True
            if(carry_on):
                carry_on=False
                cust = False
                o_stock_lu_cust_prom_th.execute("SELECT TRANS_ID,BU_CODE_LU,NOF_LINES FROM O_STOCK_LU_CUST_PROM_TH WHERE TRANS_DATE LIKE TRUNC(SYSDATE) AND BU_CODE_LU='"+cdc+"' ORDER BY TRANS_ID DESC FETCH FIRST 1 ROW ONLY")
                o_stock_lu_cust_prom_th_data = o_stock_lu_cust_prom_th.fetchall()
                if(o_stock_lu_cust_prom_th.rowcount==0):
                    try:
                        display("running IFC_STOCK_API.PROVIDE_STOCK_ISOM job for cdc "+cdc)
                        provide_stock_isom.callproc("IFC_STOCK_API.PROVIDE_STOCK_ISOM",[cdc,'CDC'])
                        wis_logic.commit()
                    except Exception as e:
                        with open("mail_body.txt", 'a') as out:
                            out.write(str(e))
                        return 0
                    cust=True
                    #StockReportISOM
                if(cust==True):
                    o_stock_lu_cust_prom_th.execute("SELECT TRANS_ID,BU_CODE_LU,NOF_LINES FROM O_STOCK_LU_CUST_PROM_TH WHERE TRANS_DATE LIKE TRUNC(SYSDATE) AND BU_CODE_LU='"+cdc+"' ORDER BY TRANS_ID DESC FETCH FIRST 1 ROW ONLY")
                    o_stock_lu_cust_prom_th_data = o_stock_lu_cust_prom_th.fetchall()
                ws_cust_th=self.wb.add_worksheet("O_STOCK_LU_CUST_PROM_TH_"+cdc)

                if(o_stock_lu_cust_prom_th.rowcount>0):
                    display("checking o_stock_lu_cust_prom_th")
                    for cust_th in  o_stock_lu_cust_prom_th_data:
                            cust_th_trans_id = str(cust_th[0])
                            cust_th_bu_code_lu=cust_th[1]
                            cust_th_nof_lines=cust_th[2]
                            ws_cust_th.write(0,0,"TRANS_ID")
                            ws_cust_th.write(0,1,"BU_CODE_LU")
                            ws_cust_th.write(0,2,"NOF_LINES")
                            ws_cust_th.write(0,3,"INS_DATE")
                            ws_cust_th.write(1,0,cust_th_trans_id)
                            ws_cust_th.write(1,1,cust_th_bu_code_lu)
                            ws_cust_th.write(1,2,cust_th_nof_lines)
                            ctrl_ins_date = transaction_ctrl.execute("SELECT INS_DATE FROM TRANSACTION_CTRL WHERE TRANS_ID ="+cust_th_trans_id).fetchone()[0].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00"                         
                            ws_cust_th.write(1,3,ctrl_ins_date)
                            if(self.onhand):
                                ws_onhand = self.wb.add_worksheet("Onhand Inventory ISOM_"+cdc)
                                onhand_r=0
                                items = []
                                item=''
                                if(var1==0):
                                    isom_r = self.wb_r.sheet_by_name("Onhand Inventory ISOM")                                   
                                    for i in range(1,isom_r.nrows):                                 
                                        items.append(isom_r.cell_value(i,0))
                                else:
                                    items=e2.get().split(",")
                                for i in range(len(items)):
                                    item+="'"+items[i]+"'"
                                    if(i<len(items)-1):
                                        item+=","                                
                                if(item=="''"):
                                    o_stock_lu_cust_prom_tl.execute("SELECT TRANS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,ITEM_QTY_AVAIL,ITEM_QTY_ISOM,UOM_CODE_QTY FROM O_STOCK_LU_CUST_PROM_TL WHERE TRANS_ID = "+cust_th_trans_id+" ORDER BY ITEM_QTY_AVAIL DESC FETCH FIRST 50 ROWS ONLY")
                                    o_stock_lu_cust_prom_tl_data=o_stock_lu_cust_prom_tl.fetchall()
                                else:
                                    o_stock_lu_cust_prom_tl.execute("SELECT TRANS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,ITEM_QTY_AVAIL,ITEM_QTY_ISOM,UOM_CODE_QTY FROM O_STOCK_LU_CUST_PROM_TL WHERE ITEM_NO IN ("+item+") AND TRANS_ID = "+cust_th_trans_id+" ORDER BY ITEM_QTY_AVAIL DESC")
                                    o_stock_lu_cust_prom_tl_data=o_stock_lu_cust_prom_tl.fetchall()

                                if(o_stock_lu_cust_prom_tl.rowcount>0):
                                    ws_onhand.write(0,0,"TRANS_ID")
                                    ws_onhand.write(0,1,"BU_CODE_LU")
                                    ws_onhand.write(0,2,"BU_TYPE_LU")
                                    ws_onhand.write(0,3,"ITEM_NO")
                                    ws_onhand.write(0,4,"ITEM_TYPE")
                                    ws_onhand.write(0,5,"ITEM_QTY_AVAIL")
                                    ws_onhand.write(0,6,"ITEM_QTY_ISOM")
                                    ws_onhand.write(0,7,"UOM_CODE_QTY")
                                    for cust_tl in o_stock_lu_cust_prom_tl_data:
                                        onhand_r+=1
                                        ws_onhand.write(onhand_r,0,cust_tl[0])
                                        ws_onhand.write(onhand_r,1,cust_tl[1])
                                        ws_onhand.write(onhand_r,2,cust_tl[2])
                                        ws_onhand.write(onhand_r,3,cust_tl[3])
                                        ws_onhand.write(onhand_r,4,cust_tl[4])
                                        ws_onhand.write(onhand_r,5,cust_tl[5])
                                        ws_onhand.write(onhand_r,6,cust_tl[6])
                                        ws_onhand.write(onhand_r,7,cust_tl[7])
                                else:
                                    ws_onhand.write(0,0,"No Entries found in O_STOCT_LU_CUST_PROM_TL_"+cdc)
                else:
                    ws_cust_th.write(0,0,"No records in o_stock_lu_cust_prom_th_"+cdc)

        provide_stock_isom.close()
        i_stock_reps.close()
        transaction_ctrl.close()
        o_stock_lu_cust_prom_th.close()
        o_stock_lu_cust_prom_tl.close()

        return 1
        
class PublishStock_OMS_GEMINI:
    def __init__(self,wb,wb_r,onhand):
        self.wb=wb
        self.wb_r=wb_r 
        self.onhand = onhand

    def processStockToOMS(self):
        global day,var1,e1,e2
        bu_code = self.wb_r.sheet_by_name("Publish Stock to OMS_GEMINI")
        publish_cdc = []
        if(var1==0):
            for i in range(1,bu_code.nrows):
                publish_cdc.append(bu_code.cell_value(i,0))
        else:
            publish_cdc=e1.get().split(",")

        provide_stock_lu= wis_logic.cursor()
        i_stock_reps=iip_obj.cursor()
        transaction_ctrl=iip_obj.cursor()
        o_stock_lu_stk_th = iip_obj.cursor()
        o_stock_lu_stk_tl = iip_obj.cursor()
        
        for cdc in publish_cdc:
            carry_on = False
            if(not self.onhand):
                display("checking i_stock_reps")
                stk_reps = self.wb.add_worksheet("I_STOCK_REPS_"+cdc)
                if(day=='Today'):
                    i_stock_reps.execute("SELECT BU_CODE_LU,TRANS_DATE,TRANS_STATUS,ERR_REASON,INS_DATE FROM I_STOCK_REPS WHERE BU_CODE_LU='"+cdc+"' AND TRANS_DATE LIKE TRUNC(sysdate) ORDER BY INS_DATE DESC")
                    i_stock_reps_data = i_stock_reps.fetchall()
                else:
                    i_stock_reps.execute("SELECT BU_CODE_LU,TRANS_DATE,TRANS_STATUS,ERR_REASON,INS_DATE FROM I_STOCK_REPS WHERE BU_CODE_LU='"+cdc+"'  ORDER BY INS_DATE DESC")
                    i_stock_reps_data=i_stock_reps.fetchall()
                if(i_stock_reps.rowcount>0):
                    for reps in i_stock_reps_data:
                        if(reps[2] in (1,7)):             
                            bu_code_lu=reps[0]
                            trans_date=reps[1].strftime("%d-%b-%Y")
                            trans_status = reps[2]
                            err_reason=reps[3 ]
                            ins_date = reps[4].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00"
                            stk_reps.write(0,0,"BU_CODE_LU")
                            stk_reps.write(0,1,"TRANS_DATE")
                            stk_reps.write(0,2,"TRANS_STATUS")
                            stk_reps.write(0,3,"ERR_REASON")
                            stk_reps.write(0,4,"INS_DATE")
                            stk_reps.write(1,0,bu_code_lu)
                            stk_reps.write(1,1,trans_date)
                            stk_reps.write(1,2,trans_status)
                            stk_reps.write(1,3,err_reason)
                            stk_reps.write(1,4,ins_date) 
                            carry_on=True
                            break
                        else:
                            stk_reps.write("Error in i_stock_reps TRANS_STATUS is "+reps[2]+" for cdc "+cdc)
                            carry_on = False
                            break
                else:
                    stk_reps.write(0,0,"No Stock Report")
            if(self.onhand):
                carry_on=True
            if(carry_on):
                carry_on=False
                stk=False
                o_stock_lu_stk_th.execute("SELECT TRANS_ID,BU_CODE_LU,NOF_LINES FROM O_STOCK_LU_STK_TH WHERE TRANS_DATE LIKE TRUNC(SYSDATE) AND BU_CODE_LU='"+cdc+"' ORDER BY TRANS_ID DESC FETCH FIRST 1 ROW ONLY")                
                o_stock_lu_stk_th_data = o_stock_lu_stk_th.fetchall()
                if(o_stock_lu_stk_th.rowcount==0):
                    try:
                        display("running IFC_STOCK_API.PROVIDE_STOCK_LU job for cdc "+cdc)
                        provide_stock_lu.callproc("IFC_STOCK_API.PROVIDE_STOCK_LU",[cdc,'CDC'])
                        wis_logic.commit()
                    except Exception as e:
                        with open("mail_body.txt", 'a') as out:
                            out.write(str(e))                   
                        return 0
                    stk=True
                
                #StockReport_OMS_GEMINI
                if(stk==True):
                    o_stock_lu_stk_th.execute("SELECT TRANS_ID,BU_CODE_LU,NOF_LINES FROM O_STOCK_LU_STK_TH WHERE TRANS_DATE LIKE TRUNC(SYSDATE) AND BU_CODE_LU='"+cdc+"' ORDER BY TRANS_ID DESC FETCH FIRST 1 ROW ONLY")                
                    o_stock_lu_stk_th_data = o_stock_lu_stk_th.fetchall()
                ws_stk_th=self.wb.add_worksheet("O_STOCK_LU_STK_TH_"+cdc)
                if(o_stock_lu_stk_th.rowcount>0):
                    display("fetching o_stock_lu_stk_th")
                    for stk_th in  o_stock_lu_stk_th_data:
                            stk_th_trans_id = str(stk_th[0])
                            stk_th_bu_code_lu=stk_th[1]
                            stk_th_nof_lines=stk_th[2]
                            ws_stk_th.write(0,0,"TRANS_ID")
                            ws_stk_th.write(0,1,"BU_CODE_LU")
                            ws_stk_th.write(0,2,"NOF_LINES")
                            ws_stk_th.write(0,3,"INS_DATE")
                            ws_stk_th.write(1,0,stk_th_trans_id)
                            ws_stk_th.write(1,1,stk_th_bu_code_lu)
                            ws_stk_th.write(1,2,stk_th_nof_lines)
                            ctrl_ins_date = transaction_ctrl.execute("SELECT INS_DATE FROM TRANSACTION_CTRL WHERE TRANS_ID ="+stk_th_trans_id).fetchone()[0].strftime("%d-%b-%y %I.%M.%S.%f %p")
                            ctrl_ins_date+=" +00:00"
                            ws_stk_th.write(1,3,ctrl_ins_date)
                            if(self.onhand):
                                ws_onhand = self.wb.add_worksheet("Onhand Inventory OMS_GEMINI_"+cdc)
                                onhand_r=0
                                items = []
                                item=''
                                if(var1==0):
                                    oms_r = self.wb_r.sheet_by_name("Onhand Inventory OMS_GEMINI")
                                    for i in range(1,oms_r.nrows):                                 
                                        items.append(oms_r.cell_value(i,0))
                                else:
                                    items=e2.get().split(",")
                                for i in range(len(items)):
                                    item+="'"+items[i]+"'"
                                    if(i<len(items)-1):
                                        item+=","
                                if(item=="''"):
                                    o_stock_lu_stk_tl.execute("SELECT TRANS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,BU_CODE_SUP,BU_TYPE_SUP,ITEM_QTY_AVAIL,UOM_CODE_QTY FROM O_STOCK_LU_STK_TL WHERE TRANS_ID = "+stk_th_trans_id+" ORDER BY ITEM_QTY_AVAIL DESC FETCH FIRST 50 ROWS ONLY")
                                    o_stock_lu_stk_tl_data=o_stock_lu_stk_tl.fetchall()
                                else:
                                    o_stock_lu_stk_tl.execute("SELECT TRANS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,BU_CODE_SUP,BU_TYPE_SUP,ITEM_QTY_AVAIL,UOM_CODE_QTY FROM O_STOCK_LU_STK_TL WHERE ITEM_NO IN ("+item+") AND TRANS_ID = "+stk_th_trans_id+" ORDER BY ITEM_QTY_AVAIL DESC")
                                    o_stock_lu_stk_tl_data=o_stock_lu_stk_tl.fetchall()
                                if(o_stock_lu_stk_tl.rowcount>0):
                                    ws_onhand.write(0,0,"TRANS_ID")
                                    ws_onhand.write(0,1,"BU_CODE_LU")
                                    ws_onhand.write(0,2,"BU_TYPE_LU")
                                    ws_onhand.write(0,3,"ITEM_NO")
                                    ws_onhand.write(0,4,"ITEM_TYPE")
                                    ws_onhand.write(0,5,"BU_CODE_SUP")
                                    ws_onhand.write(0,6,"BU_TYPE_SUP")
                                    ws_onhand.write(0,7,"ITEM_QTY_AVAIL")
                                    ws_onhand.write(0,8,"UOM_CODE_QTY")
                                    for stk_tl in o_stock_lu_stk_tl_data:
                                        onhand_r+=1
                                        ws_onhand.write(onhand_r,0,stk_tl[0])
                                        ws_onhand.write(onhand_r,1,stk_tl[1])
                                        ws_onhand.write(onhand_r,2,stk_tl[2])
                                        ws_onhand.write(onhand_r,3,stk_tl[3])
                                        ws_onhand.write(onhand_r,4,stk_tl[4])
                                        ws_onhand.write(onhand_r,5,stk_tl[5])
                                        ws_onhand.write(onhand_r,6,stk_tl[6])
                                        ws_onhand.write(onhand_r,7,stk_tl[7])
                                        ws_onhand.write(onhand_r,8,stk_tl[8])
                                else:
                                    ws_onhand.write(0,0,"No Entries found in O_STOCT_LU_STK_TL_"+cdc)                                                                 
                else:
                    ws_stk_th.write(0,0,"No records in o_stock_lu_stk_th_"+cdc)

        provide_stock_lu.close()
        i_stock_reps.close()
        transaction_ctrl.close()
        o_stock_lu_stk_th.close()
        o_stock_lu_stk_tl.close()

        return 1

class Cancellation:

    def __init__(self,wb,wb_r):
        self.wb=wb
        self.wb_r=wb_r

    def processCancellation(self):
        global var1,e1,e2
        cancel_ord_r = self.wb_r.sheet_by_name("Cancellation")
        ord=''
        cancel_ord=[]
        wrk_ord_ref_dict={}
        carry_on = False
        if(var1==0):
            for i in range(1,cancel_ord_r.nrows):
                cancel_ord.append(cancel_ord_r.cell_value(i,0))
        else:
            cancel_ord=e1.get().split(",")
        for i in range(len(cancel_ord)):
            ord+="'"+cancel_ord[i]+"'"
            if(i<len(cancel_ord)-1):
                ord+=","
        i_orders=iip_obj.cursor()
        i_order_changes_dlv = iip_obj.cursor()
        order_lines =wis_data.cursor()
        order_line_changes = wis_data.cursor()
        order_changes = wis_data.cursor()
        process_order_changes_dlv=wis_logic.cursor()

        display("fetching cancellation details")
      
        i_orders.execute("SELECT BU_CODE_LU,WRK_ORD_REF,ORG_ORD_REF,TRANS_STATUS,ORD_ID_REF_SALES FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+")")
        i_orders_data = i_orders.fetchall()
        ws_ord=self.wb.add_worksheet("I_ORDERS")
        ord_r=0
        if(i_orders.rowcount>0):
            ws_ord.write(0,0,"BU_CODE_LU")
            ws_ord.write(0,1,"WRK_ORD_REF")
            ws_ord.write(0,2,"ORG_ORD_REF")
            ws_ord.write(0,3,"TRANS_STATUS")
            ws_ord.write(0,4,"ORD_ID_REF_SALES")
            for order in i_orders_data:
                ord_r+=1
                ws_ord.write(ord_r,0,order[0])
                ws_ord.write(ord_r,1,order[1])
                ws_ord.write(ord_r,2,order[2])
                ws_ord.write(ord_r,3,order[3])
                ws_ord.write(ord_r,4,order[4])
                    
            i_order_changes_dlv.execute("SELECT BU_CODE_LU,ORD_SRC,ORG_ORD_REF,CHG_ACTION,TRANS_STATUS,INS_DATE FROM I_ORDER_CHANGES_DLV WHERE ORG_ORD_REF IN (SELECT ORG_ORD_REF FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+"))")
            i_order_changes_dlv_data = i_order_changes_dlv.fetchall()
            if(i_order_changes_dlv.rowcount==0):
                i_order_changes_dlv.execute("SELECT BU_CODE_LU,ORD_SRC,ORG_ORD_REF,CHG_ACTION,TRANS_STATUS,INS_DATE FROM I_ORDER_CHANGES_DLV WHERE ORG_ORD_REF IN (SELECT WRK_ORD_REF FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+"))")
                i_order_changes_dlv_data = i_order_changes_dlv.fetchall()
            ws_dlv = self.wb.add_worksheet("ORDER_LINE_CHANGES_DLV")
            dlv_r = 0
            if(i_order_changes_dlv.rowcount>0):
                ord_chg = False
                for dlv in i_order_changes_dlv_data:
                    if(dlv[4]==0):
                        try:
                            display("running ORDER_MODIFICATION_API.PROCESS_ORDER_CHANGES_DLV job")
                            process_order_changes_dlv.callproc("ORDER_MODIFICATION_API.PROCESS_ORDER_CHANGES_DLV")
                            wis_logic.commit()
                            ord_chg=True
                        except Exception as e:
                            with open("log.txt", 'a') as out:
                                out.write(str(e))                   
                            return 0
                
                ws_dlv.write(0,0,"BU_CODE_LU")
                ws_dlv.write(0,1,"ORD_SRC")
                ws_dlv.write(0,2,"ORD_ORD_REF")
                ws_dlv.write(0,3,"CHG_ACTION")
                ws_dlv.write(0,4,"TRANS_STATUS")
                ws_dlv.write(0,5,"INS_DATE")

                if(ord_chg):
                    i_order_changes_dlv.execute("SELECT BU_CODE_LU,ORD_SRC,ORG_ORD_REF,CHG_ACTION,TRANS_STATUS,INS_DATE FROM I_ORDER_CHANGES_DLV WHERE ORG_ORD_REF IN (SELECT ORG_ORD_REF FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+"))")
                    i_order_changes_dlv_data = i_order_changes_dlv.fetchall()
                    if(i_order_changes_dlv.rowcount==0):
                        i_order_changes_dlv.execute("SELECT BU_CODE_LU,ORD_SRC,ORG_ORD_REF,CHG_ACTION,TRANS_STATUS,INS_DATE FROM I_ORDER_CHANGES_DLV WHERE ORG_ORD_REF IN (SELECT WRK_ORD_REF FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+"))")
                        i_order_changes_dlv_data = i_order_changes_dlv.fetchall()

                for dlv in i_order_changes_dlv_data:
                    dlv_r+=1
                    ws_dlv.write(dlv_r,0,dlv[0])
                    ws_dlv.write(dlv_r,1,dlv[1])
                    ws_dlv.write(dlv_r,2,dlv[2])
                    ws_dlv.write(dlv_r,3,dlv[3])
                    ws_dlv.write(dlv_r,4,dlv[4])
                    ws_dlv.write(dlv_r,5,dlv[5].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                    if(dlv[4]==7):
                        carry_on=True
            else:
                ws_dlv.write(0,0,"Cancellation is not received in CWIS")

            if(carry_on):
                carry_on = False
                order_lines.execute("SELECT ordl.ORD_ID,ordl.ITEM_NO,ordl.ORDL_STATUS,ordl.INS_DATE,ordl.ORD_ID_IN_TK,ordl.CANCEL_SRC,ord.WRK_ORD_REF,ord.BU_CODE_LU,ord.ORG_ORD_REF,ord.ORD_ID_REF_SALES FROM ORDER_LINES ordl,I_ORDERS ord WHERE ord.ORD_ID_TK=ordl.ORD_ID_IN_TK and ord.ORD_ID_REF_SALES IN ("+ord+")")
                order_lines_data = order_lines.fetchall()
                ws_ordl=self.wb.add_worksheet("ORDER_LINES")
                ordl_r=0
                if(order_lines.rowcount>0):
                    ws_ordl.write(0,0,"ORD_ID")
                    ws_ordl.write(0,1,"ITEM_NO")
                    ws_ordl.write(0,2,"ORDL_STATUS")
                    ws_ordl.write(0,3,"INS_DATE")
                    ws_ordl.write(0,4,"ORD_ID_IN_TK")
                    ws_ordl.write(0,5,"CANCEL_SRC")
                    ws_ordl.write(0,6,"WRK_ORD_REF")
                    ws_ordl.write(0,7,"BU_CODE_LU")
                    ws_ordl.write(0,8,"ORG_ORD_REF")
                    ws_ordl.write(0,9,"ORD_ID_REF_SALES")
                    for ordl in order_lines_data:
                        ordl_r+=1
                        wrk_ord_ref_dict[ordl[0]]=ordl[6]
                        ws_ordl.write(ordl_r,0,ordl[0])
                        ws_ordl.write(ordl_r,1,ordl[1])
                        ws_ordl.write(ordl_r,2,ordl[2])
                        ws_ordl.write(ordl_r,3,ordl[3].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                        ws_ordl.write(ordl_r,4,ordl[4])
                        ws_ordl.write(ordl_r,5,ordl[5])
                        ws_ordl.write(ordl_r,6,ordl[6])
                        ws_ordl.write(ordl_r,7,ordl[7])
                        ws_ordl.write(ordl_r,8,ordl[8])
                        ws_ordl.write(ordl_r,9,ordl[9])
                        if(ordl[2] in (40,30,80,90)):
                            carry_on = True
                else:
                    ws_ordl.write(0,0,"No Entries in ORDER LINES")
            if(carry_on):
                order_line_changes.execute("SELECT CHG_SRC,CHG_ACTION,CHG_STATUS,ORD_ID,INS_DATE,ITEM_NO,SND_STATUS FROM ORDER_LINE_CHANGES  WHERE ORD_ID IN (SELECT ORD_ID FROM ORDER_LINES WHERE ORD_ID_IN_TK IN (SELECT ORD_ID_TK FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+")))")
                order_line_changes_data = order_line_changes.fetchall()
                
                order_changes.execute("SELECT CHG_SRC,CHG_ACTION,CHG_STATUS,ORD_ID,INS_DATE FROM ORDER_CHANGES  WHERE ORD_ID IN (SELECT ORD_ID FROM ORDER_LINES WHERE ORD_ID_IN_TK IN (SELECT ORD_ID_TK FROM I_ORDERS WHERE ORD_ID_REF_SALES IN ("+ord+")))")
                order_changes_data = order_changes.fetchall()

                ws_chg = self.wb.add_worksheet("ORDER_LINE_CHANGES")
                ws_log=self.wb.add_worksheet("XML_MESSAGE_LOG")
                chg_r=0
                chg_status=False
                
                if(order_line_changes.rowcount>0):
                    chg_status=True
                    ws_chg.write(0,0,"CHG_SRC")
                    ws_chg.write(0,1,"CHG_ACTION")
                    ws_chg.write(0,2,"CHG_STATUS")
                    ws_chg.write(0,3,"ORD_ID")
                    ws_chg.write(0,4,"INS_DATE")
                    ws_chg.write(0,5,"ITEM_NO")
                    ws_chg.write(0,6,"SND_STATUS")                   
                        
                    for chg in order_line_changes_data:
                        chg_r+=1
                        exe=[]
                        wrk=[]
                        ws_chg.write(chg_r,0,chg[0])
                        ws_chg.write(chg_r,1,chg[1])
                        ws_chg.write(chg_r,2,chg[2])
                        ws_chg.write(chg_r,3,chg[3])
                        ws_chg.write(chg_r,4,chg[4].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                        ws_chg.write(chg_r,5,chg[5])
                        ws_chg.write(chg_r,6,chg[6])
                        exe.append(str(chg[3]))
                        wrk.append(str(wrk_ord_ref_dict[chg[3]]))
                        if(chg[2]=='RECEIVED'):
                            display("check for ModifyExecutableOrderLineMsg in logs")
                            checkLogs(ws_log,chg_r,'ModifyExecutableOrderLineMsg','ExecutableOrderId',exe,'Cacellation is sent to ASTRO')                                
                        if(chg[2]=='EXECUTED'):
                            display("check for ConfirmReleaseOrderModificationMsg in logs")
                            checkLogs(ws_log,chg_r,'ConfirmReleaseOrderModificationMsg','WorkOrderReference',wrk,'Cancellation Confirmation is sent to ISOM')
                                                                                   
                                                                                                                     
                if(order_changes.rowcount>0):
                    chg_status=True
                    ws_chg.write(0,0,"CHG_SRC")
                    ws_chg.write(0,1,"CHG_ACTION")
                    ws_chg.write(0,2,"CHG_STATUS")
                    ws_chg.write(0,3,"ORD_ID")
                    ws_chg.write(0,4,"INS_DATE")
                    ws_chg.write(0,5,"ITEM_NO")
                    ws_chg.write(0,6,"SND_STATUS")
                    
                    for chg in order_changes_data:
                        chg_r+=1
                        exe=[]
                        wrk=[]
                        ws_chg.write(chg_r,0,chg[0])
                        ws_chg.write(chg_r,1,chg[1])
                        ws_chg.write(chg_r,2,chg[2])
                        ws_chg.write(chg_r,3,chg[3])
                        ws_chg.write(chg_r,4,chg[4].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                        ws_chg.write(chg_r,5,'')
                        ws_chg.write(chg_r,6,'')
                        exe.append(str(chg[3]))
                        wrk.append(str(wrk_ord_ref_dict[chg[3]]))
                        if(chg[2]=='RECEIVED'):
                            display("check for ModifyExecutableOrderLineMsg in logs")
                            checkLogs(ws_log,chg_r,'ModifyExecutableOrderLineMsg','ExecutableOrderId',exe,'Cacellation is sent to ASTRO')                                
                        if(chg[2]=='EXECUTED'):
                            display("check for ConfirmReleaseOrderModificationMsg in logs")
                            checkLogs(ws_log,chg_r,'ConfirmReleaseOrderModificationMsg','WorkOrderReference',wrk,'Cancellation Confirmation is sent to ISOM')
               
                if(chg_status==False):
                    ws_chg.write("No Entries in ORDER_CHANGES/ORDER_LINE_CHANGES")
        else:
            ws_ord.write(0,0,"No entries in I_ORDERS")
            
        i_orders.close()
        i_order_changes_dlv.close()
        order_lines.close()
        order_line_changes.close()
        order_changes.close()
        process_order_changes_dlv.close()

        return 1

class StockAdjustments:

    def __init__(self,wb,wb_r):
        self.wb=wb
        self.wb_r = wb_r

    def process_adjustments(self):
        global var1,e1,e2,day
        stock_adjs = self.wb_r.sheet_by_name("Stock adjustment")
        cdc = ''
        items = []
        cdcs=[]
        articles = ''
        carry_on = False
        latest_date = ''
        if(var1==0):
            cdc = stock_adjs.cell_value(1,0)
            for i in range(1,stock_adjs.nrows):
                items.append(stock_adjs.cell_value(i,1))
        else:
            cdc=e1.get()
            items = e2.get().split(",")
        for i in range(len(items)):
            articles+="'"+items[i]+"'"
            if(i<len(items)-1):
                articles+=","

        cdcs.append(cdc)

        i_stock_adjs = iip_obj.cursor()
        o_stock_adjs_isom = iip_obj.cursor()
        
        display("fetching stock adjustments")
        if(articles=="''"):
            if(day=='Today'):
                i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE TRUNC(SYSDATE) AND BU_CODE_LU ="+cdc)
                i_stock_adjs_data = i_stock_adjs.fetchall()
            else:
                i_stock_adjs.execute("SELECT TRUNC(INS_DATE) FROM I_STOCK_ADJS WHERE  BU_CODE_LU ="+cdc+" ORDER BY INS_DATE DESC FETCH FIRST 1 ROW ONLY")
                i_stk_adjs_data=i_stock_adjs.fetchall()
                if(i_stock_adjs.rowcount>0):
                    latest_date=i_stk_adjs_data[0][0].fetchone().strftime("%d-%b-%y").upper()
                    i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE '"+latest_date+"' AND BU_CODE_LU ="+cdc)
                    i_stock_adjs_data = i_stock_adjs.fetchall()
        else:
            if(day=='Today'):
                i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE TRUNC(SYSDATE) AND BU_CODE_LU ="+cdc+" AND ITEM_NO IN ("+articles+")")
                i_stock_adjs_data = i_stock_adjs.fetchall()
            else:               
                i_stock_adjs.execute("SELECT TRUNC(INS_DATE) FROM I_STOCK_ADJS WHERE  BU_CODE_LU ="+cdc+" AND ITEM_NO IN ("+articles+") ORDER BY INS_DATE DESC FETCH FIRST 1 ROW ONLY")
                i_stk_adjs_data=i_stock_adjs.fetchall()
                if(i_stock_adjs.rowcount>0):
                    latest_date=i_stk_adjs_data[0][0].strftime("%d-%b-%y").upper()
                    i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE '"+latest_date+"' AND BU_CODE_LU ="+cdc+" AND ITEM_NO IN ("+articles+")")
                    i_stock_adjs_data = i_stock_adjs.fetchall()
                       
        ws_adjs = self.wb.add_worksheet("I_STOCK_ADJS")
        adjs_r = 0
        if(i_stock_adjs.rowcount>0):
            ws_adjs.write(0,0,"BU_CODE_LU")
            ws_adjs.write(0,1,"LOG_ID")
            ws_adjs.write(0,2,"TRANS_TYPE")
            ws_adjs.write(0,3,"TRANS_DATE")
            ws_adjs.write(0,4,"ADJUST_QTY")
            ws_adjs.write(0,5,"ITEM_NO")
            ws_adjs.write(0,6,"BU_CODE_SUP")
            ws_adjs.write(0,7,"CSM_NO_OUT")
            ws_adjs.write(0,8,"ITEM_QTY")
            ws_adjs.write(0,9,"CSM_NO_IN")
            ws_adjs.write(0,10,"TRANS_STATUS")
            ws_adjs.write(0,11,"INS_DATE")
            ws_adjs.write(0,12,"ERR_REASON")
            for adjs in i_stock_adjs_data:
                if(adjs[10]==0):
                    display("running IFC_STOCK_ADJ_API.integrate_data job...")
                    try:
                        integrate_data=wis_logic.cursor()
                        integrate_data.callproc("IFC_STOCK_ADJ_API.integrate_data",[cdc,'CDC'])
                        wis_logic.commit()
                    except Exception as e:
                        with open("log.txt", 'a') as out:
                            out.write(str(e))
                        return 0
                    if(articles=="''"):
                        if(day=='Today'):
                            i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE TRUNC(SYSDATE) AND BU_CODE_LU ="+cdc)
                            i_stock_adjs_data = i_stock_adjs.fetchall()
                        else:
                            i_stock_adjs.execute("SELECT TRUNC(INS_DATE) FROM I_STOCK_ADJS WHERE  BU_CODE_LU ="+cdc+" ORDER BY INS_DATE DESC FETCH FIRST 1 ROW ONLY")
                            i_stk_adjs_data=i_stock_adjs.fetchall()
                            if(i_stock_adjs.rowcount>0):
                                latest_date=i_stk_adjs_data[0][0].fetchone().strftime("%d-%b-%y").upper()
                                i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE '"+latest_date+"' AND BU_CODE_LU ="+cdc)
                                i_stock_adjs_data = i_stock_adjs.fetchall()
                    else:
                        if(day=='Today'):
                            i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE TRUNC(SYSDATE) AND BU_CODE_LU ="+cdc+" AND ITEM_NO IN ("+articles+")")
                            i_stock_adjs_data = i_stock_adjs.fetchall()
                        else:               
                            i_stock_adjs.execute("SELECT TRUNC(INS_DATE) FROM I_STOCK_ADJS WHERE  BU_CODE_LU ="+cdc+" AND ITEM_NO IN ("+articles+") ORDER BY INS_DATE DESC FETCH FIRST 1 ROW ONLY")
                            i_stk_adjs_data=i_stock_adjs.fetchall()
                            if(i_stock_adjs.rowcount>0):
                                latest_date=i_stk_adjs_data[0][0].strftime("%d-%b-%y").upper()
                                i_stock_adjs.execute("SELECT BU_CODE_LU,LOG_ID,TRANS_TYPE,TRANS_DATE,ADJUST_QTY,ITEM_NO,BU_CODE_SUP,CSM_NO_OUT,ITEM_QTY,CSM_NO_IN,TRANS_STATUS,INS_DATE,ERR_REASON FROM I_STOCK_ADJS WHERE TRUNC(INS_DATE) LIKE '"+latest_date+"' AND BU_CODE_LU ="+cdc+" AND ITEM_NO IN ("+articles+")")
                                i_stock_adjs_data = i_stock_adjs.fetchall() 
                    break
            for adjs in i_stock_adjs_data:
                adjs_r+=1    
                ws_adjs.write(adjs_r,0,adjs[0])
                ws_adjs.write(adjs_r,1,adjs[1])
                ws_adjs.write(adjs_r,2,adjs[2])
                ws_adjs.write(adjs_r,3,adjs[3].strftime("%d-%b-%y"))
                ws_adjs.write(adjs_r,4,adjs[4])
                ws_adjs.write(adjs_r,5,adjs[5])
                ws_adjs.write(adjs_r,6,adjs[6])
                ws_adjs.write(adjs_r,7,adjs[7])
                ws_adjs.write(adjs_r,8,adjs[8])
                ws_adjs.write(adjs_r,9,adjs[9])
                ws_adjs.write(adjs_r,10,adjs[10])
                ws_adjs.write(adjs_r,11,adjs[11].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                ws_adjs.write(adjs_r,12,adjs[12])
                if(adjs[10]==7):
                    carry_on=True
        else:
            display("No adjustments found :( checking logs!..")
            ws_log=self.wb.add_worksheet("XML_MESSAGE_LOG")
            if(len(items)>0):
                log_status=checkLogs(ws_log,1,'SyncLUStockAdjustmentMsg','ItemNo',items,'Stock Adjustments are stuck in XML_MESSAGE_LOG')
            else:
                log_status=checkLogs(ws_log,1,'SyncLUStockAdjustmentMsg','BusinessUnitCodeLU',cdcs,'Stock Adjustments are stuck in XML_MESSAGE_LOG')
            if(not log_status):
                ws_adjs.write(0,0,"No Adjustments received")
        if(carry_on):
            if(day=='Today'):
                o_stock_adjs_isom.execute("SELECT TRANS_DATE_TIME,SYNC_LU_STOCK_STATUS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,BU_CODE_SUP,BU_TYPE_SUP,ISOM_STOCK_IN_LU_CHANGE,ISOM_TOTAL_BLOCK_CHANGE,UOM_CODE_QTY,CSM_NO,IS_RETURN,IS_INDELIVERY,TRANS_STATUS,INS_DATE FROM O_STOCK_ADJS_ISOM WHERE TRUNC(TRANS_DATE_TIME) LIKE TRUNC(SYSDATE) AND BU_CODE_LU = "+cdc+" AND ITEM_NO IN ("+articles+")")
                o_stock_adjs_isom_data = o_stock_adjs_isom.fetchall()
            else:
                o_stock_adjs_isom.execute("SELECT TRANS_DATE_TIME,SYNC_LU_STOCK_STATUS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,BU_CODE_SUP,BU_TYPE_SUP,ISOM_STOCK_IN_LU_CHANGE,ISOM_TOTAL_BLOCK_CHANGE,UOM_CODE_QTY,CSM_NO,IS_RETURN,IS_INDELIVERY,TRANS_STATUS,INS_DATE FROM O_STOCK_ADJS_ISOM WHERE TRUNC(TRANS_DATE_TIME) LIKE '"+latest_date+"' AND BU_CODE_LU = "+cdc+" AND ITEM_NO IN ("+articles+")")
                o_stock_adjs_isom_data = o_stock_adjs_isom.fetchall()

            ws_isom = self.wb.add_worksheet("O_STOCK_ADJS_ISOM")
            isom_r = 0
            if(o_stock_adjs_isom.rowcount>0):
                ws_isom.write(0,0,"TRANS_DATE_TIME")
                ws_isom.write(0,1,"SYNC_LU_STOCK_STATUS_ID")
                ws_isom.write(0,2,"BU_CODE_LU")
                ws_isom.write(0,3,"BU_TYPE_LU")
                ws_isom.write(0,4,"ITEM_NO")
                ws_isom.write(0,5,"ITEM_TYPE")
                ws_isom.write(0,6,"BU_CODE_SUP")
                ws_isom.write(0,7,"BU_TYPE_SUP")
                ws_isom.write(0,8,"ISOM_STOCK_IN_LU_CHANGE")
                ws_isom.write(0,9,"ISOM_TOTAL_BLOCK_CHANGE")
                ws_isom.write(0,10,"UOM_CODE_QTY")
                ws_isom.write(0,11,"CSM_NO")
                ws_isom.write(0,12,"IS_RETURN")
                ws_isom.write(0,13,"IS_INDELIVERY")
                ws_isom.write(0,14,"TRANS_STATUS")
                ws_isom.write(0,15,"INS_DATE")
                for isom in o_stock_adjs_isom_data:
                    if(isom[14]==0):
                        display("running IFC_STOCK_API.REQUEST_STOCK_ISOM job..")
                        try:
                            request_stock_isom=iip_obj.cursor()
                            request_stock_isom.callproc("IFC_STOCK_API.REQUEST_STOCK_ISOM",[0])
                            wis_logic.commit()
                        except Exception as e:
                            with open("log.txt", 'a') as out:
                                out.write(str(e))
                            return 0
                        if(day=='Today'):
                            o_stock_adjs_isom.execute("SELECT TRANS_DATE_TIME,SYNC_LU_STOCK_STATUS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,BU_CODE_SUP,BU_TYPE_SUP,ISOM_STOCK_IN_LU_CHANGE,ISOM_TOTAL_BLOCK_CHANGE,UOM_CODE_QTY,CSM_NO,IS_RETURN,IS_INDELIVERY,TRANS_STATUS,INS_DATE FROM O_STOCK_ADJS_ISOM WHERE TRUNC(TRANS_DATE_TIME) LIKE TRUNC(SYSDATE) AND BU_CODE_LU = "+cdc+" AND ITEM_NO IN ("+articles+")")
                            o_stock_adjs_isom_data = o_stock_adjs_isom.fetchall()
                        else:
                            o_stock_adjs_isom.execute("SELECT TRANS_DATE_TIME,SYNC_LU_STOCK_STATUS_ID,BU_CODE_LU,BU_TYPE_LU,ITEM_NO,ITEM_TYPE,BU_CODE_SUP,BU_TYPE_SUP,ISOM_STOCK_IN_LU_CHANGE,ISOM_TOTAL_BLOCK_CHANGE,UOM_CODE_QTY,CSM_NO,IS_RETURN,IS_INDELIVERY,TRANS_STATUS,INS_DATE FROM O_STOCK_ADJS_ISOM WHERE TRUNC(TRANS_DATE_TIME) LIKE '"+latest_date+"' AND BU_CODE_LU = "+cdc+" AND ITEM_NO IN ("+articles+")")
                            o_stock_adjs_isom_data = o_stock_adjs_isom.fetchall()
                        break
                        

                for isom in o_stock_adjs_isom_data:
                    isom_r+=1
                    ws_isom.write(isom_r,0,isom[0].strftime("%d-%b-%y"))
                    ws_isom.write(isom_r,1,isom[1])
                    ws_isom.write(isom_r,2,isom[2])
                    ws_isom.write(isom_r,3,isom[3])
                    ws_isom.write(isom_r,4,isom[4])
                    ws_isom.write(isom_r,5,isom[5])
                    ws_isom.write(isom_r,6,isom[6])
                    ws_isom.write(isom_r,7,isom[7])
                    ws_isom.write(isom_r,8,isom[8])
                    ws_isom.write(isom_r,9,isom[9])
                    ws_isom.write(isom_r,10,isom[10])
                    ws_isom.write(isom_r,11,isom[11])
                    ws_isom.write(isom_r,12,isom[12])
                    ws_isom.write(isom_r,13,isom[13])
                    ws_isom.write(isom_r,14,isom[14])
                    ws_isom.write(isom_r,15,isom[15].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
            else:
                ws_isom.write(0,0,"No Record's in O_STOCK_ADJS_ISOM")

        i_stock_adjs.close()
        o_stock_adjs_isom.close()   
        return 1

class StockReservation:
#Stock reservation
    def __init__(self,wb,wb_r):
        self.wb=wb
        self.wb_r=wb_r

    def processStockReservation(self):
        global var1,e1,e2
        start_time=''
        end_time=''
        #Reading orders from input.xlsx
        ord=''
        orders = self.wb_r.sheet_by_name("Stock reservation")
        sales_orders = []
        if(var1==0):
            for i in range(1,orders.nrows):
                sales_orders.append(orders.cell_value(i,0))
        else:
            sales_orders = e1.get().split(",")

        for index in range(len(sales_orders)):
                ord+="'"+sales_orders[index]+"'"
                if(index<len(sales_orders)-1):
                    ord+=","
            
        i_stock_resv_reqs=iip_obj.cursor()
        i_stock_resv_req_lines=iip_obj.cursor()
        o_stock_updates=iip_obj.cursor()
        stock_resv_reqs=wis_data.cursor()
        stock_resv_req_lines=wis_data.cursor()
        stock_resvs=wis_data.cursor()
        process_stock_reservation = wis_logic.cursor()
        
        display("Fetching Stock Reservations")
        carry_on = False
        i_stock_resv_reqs.execute("SELECT TRANS_ID,RESV_REQ_ID,ORD_ID_REF_SALES,TRANSACTION_DATE,TRANS_STATUS,ERR_REASON FROM I_STOCK_RESV_REQS WHERE ORD_ID_REF_SALES IN ("+ord+") ORDER BY TRANS_ID DESC")
        i_stock_resv_reqs_data = i_stock_resv_reqs.fetchall()
        resv = self.wb.add_worksheet("I_STOCK_RESV_REQS")
        stk_reqs = 0                        
        if(i_stock_resv_reqs.rowcount>0):
            resv.write(0,0,"TRANS_ID")
            resv.write(0,1,"RESV_REQ_ID")
            resv.write(0,2,"ORD_ID_REF_SALES")
            resv.write(0,3,"TRASACTION_DATE")
            resv.write(0,4,"TRANS_STATUS")
            resv.write(0,5,"ERR_REASON")
            resv_chg= False
            for res_p in i_stock_resv_reqs_data:
                if(res_p[4]==0):
                    display("running IFC_STOCK_RESERVATIONS_API.PROCESS_STOCK_RESERVATION job")
                    print(res_p[0])
                    process_stock_reservation.callproc("IFC_STOCK_RESERVATIONS_API.PROCESS_STOCK_RESERVATION",[res_p[0]])
                    wis_logic.commit()
                    resv_chg = True

            if(resv_chg):
                i_stock_resv_reqs.execute("SELECT TRANS_ID,RESV_REQ_ID,ORD_ID_REF_SALES,TRANSACTION_DATE,TRANS_STATUS,ERR_REASON FROM I_STOCK_RESV_REQS WHERE ORD_ID_REF_SALES IN ("+ord+") ORDER BY TRANS_ID DESC")
                i_stock_resv_reqs_data = i_stock_resv_reqs.fetchall()

            for reqs in i_stock_resv_reqs_data:
                display("Fetching Reservation Details")
                if(stk_reqs==0):
                    start_time = (reqs[3]+timedelta(minutes=2)).strftime("%d-%b-%y %I.%M.%S.%f %p").upper()+" +00:00"
                if(stk_reqs==(i_stock_resv_reqs.rowcount-1)):
                    end_time = (reqs[3]-timedelta(minutes=2)).strftime("%d-%b-%y %I.%M.%S.%f %p").upper()+" +00:00"
                stk_reqs+=1
                resv.write(stk_reqs,0,reqs[0])
                resv.write(stk_reqs,1,reqs[1])
                resv.write(stk_reqs,2,reqs[2])
                resv.write(stk_reqs,3,reqs[3].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                resv.write(stk_reqs,4,reqs[4])
                resv.write(stk_reqs,5,reqs[5])
                if(reqs[4]==7):
                    carry_on=True
        else:
            resv.write(0,0,"No Entries in I_STOCK_RESV_REQS")
        
        i_stock_resv_req_lines.execute("SELECT TRANS_ID,RESV_REQ_LINE_ID,BU_CODE_LU,ITEM_NO,ITEM_QTY_REQ FROM I_STOCK_RESV_REQ_LINES WHERE TRANS_ID IN (SELECT TRANS_ID FROM I_STOCK_RESV_REQS WHERE ORD_ID_REF_SALES IN ("+ord+")) ORDER BY TRANS_ID DESC")
        i_stock_resv_req_lines_data = i_stock_resv_req_lines.fetchall()
        req = self.wb.add_worksheet("I_STOCK_RESV_REQ_LINES")
        stk_reql=0
        if(i_stock_resv_req_lines.rowcount>0):
            req.write(0,0,"TRANS_ID")
            req.write(0,1,"RESV_REQ_LINE_ID")
            req.write(0,2,"BU_CODE_LU")
            req.write(0,3,"ITEM_NO")
            req.write(0,4,"ITEM_QTY_REQ")
            for i_req in i_stock_resv_req_lines_data:
                stk_reql+=1
                req.write(stk_reql,0,i_req[0])
                req.write(stk_reql,1,i_req[1])
                req.write(stk_reql,2,i_req[2])
                req.write(stk_reql,3,i_req[3])
                req.write(stk_reql,4,i_req[4])              
        else:
            req.write(0,0,"No Records in I_STOCK_RESV_REQ_LINES")
        if(carry_on):
            stock_resv_reqs.execute("SELECT RESV_ID,ORD_ID_REF_SALES,RESV_REQ_ID,TRANSACTION_DATE FROM STOCK_RESV_REQS WHERE ORD_ID_REF_SALES IN ("+ord+") ORDER BY TRANSACTION_DATE DESC")
            stock_resv_reqs_data = stock_resv_reqs.fetchall()
            stk_res = self.wb.add_worksheet("STOCK_RESV_REQS")
            res=0
            if(stock_resv_reqs.rowcount>0):
                stk_res.write(0,0,"RESV_ID")
                stk_res.write(0,1,"ORD_ID_REF_SALES")
                stk_res.write(0,2,"RESV_REQ_ID")
                stk_res.write(0,3,"TRANSACTION_DATE")
                for req_data in stock_resv_reqs_data:
                    res+=1
                    stk_res.write(res,0,req_data[0])
                    stk_res.write(res,1,req_data[1])
                    stk_res.write(res,2,req_data[2])
                    stk_res.write(res,3,req_data[3].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
            else:
                stk_res.write(0,0,"No Entries in STOCK_RESV_REQS")

            stock_resv_req_lines.execute("SELECT RESV_ID,RESV_REQ_LINE_ID,LU_TK,ITEM_NO,ITEM_QTY_REQ,INS_DATE FROM STOCK_RESV_REQ_LINES WHERE RESV_ID IN (SELECT RESV_ID FROM STOCK_RESV_REQS WHERE ORD_ID_REF_SALES IN ("+ord+")) ORDER BY INS_DATE DESC")
            stock_resv_req_lines_data = stock_resv_req_lines.fetchall()
            reql=self.wb.add_worksheet("STOCK_RESV_REQ_LINES")
            r_reql=0
            if(stock_resv_req_lines.rowcount>0):
                reql.write(0,0,"RESV_ID")
                reql.write(0,1,"RESV_REQ_LINE_ID")
                reql.write(0,2,"LU_TK")
                reql.write(0,3,"ITEM_NO")
                reql.write(0,4,"ITEM_QTY_REQ")
                reql.write(0,5,"INS_DATE")
                for reql_data in stock_resv_req_lines_data:
                    r_reql+=1
                    reql.write(r_reql,0,reql_data[0])
                    reql.write(r_reql,1,reql_data[1])
                    reql.write(r_reql,2,reql_data[2])
                    reql.write(r_reql,3,reql_data[3])
                    reql.write(r_reql,4,reql_data[4])
                    reql.write(r_reql,5,reql_data[5].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
            else:
                reql.write(0,0,"No entries in STOCK_RESV_REQ_LINES")

            stock_resvs.execute("SELECT RESV_ID,LU_TK,ITEM_NO,BU_CODE_SUP,RESV_REQ_LINE_ID,ITEM_QTY_RESV_ORG,ITEM_QTY_RESV,INS_DATE FROM STOCK_RESVS WHERE RESV_ID IN (SELECT RESV_ID FROM STOCK_RESV_REQS WHERE ORD_ID_REF_SALES IN ("+ord+")) ORDER BY INS_DATE DESC")
            stock_resvs_data=stock_resvs.fetchall()
            resvs=self.wb.add_worksheet("STOCK_RESVS")
            r_resvs=0
            if(stock_resvs.rowcount>0):
                resvs.write(0,0,"RESV_ID")
                resvs.write(0,1,"LU_TK")
                resvs.write(0,2,"ITEM_NO")
                resvs.write(0,3,"BU_CODE_SUP")
                resvs.write(0,4,"RESV_REQ_LINE_ID")
                resvs.write(0,5,"ITEM_QTY_RESV_ORG")
                resvs.write(0,6,"ITEM_QTY_RESV")
                resvs.write(0,7,"INS_DATE")
                for resvs_data in stock_resvs_data:
                    r_resvs+=1
                    resvs.write(r_resvs,0,resvs_data[0])
                    resvs.write(r_resvs,1,resvs_data[1])
                    resvs.write(r_resvs,2,resvs_data[2])
                    resvs.write(r_resvs,3,resvs_data[3])
                    resvs.write(r_resvs,4,resvs_data[4])
                    resvs.write(r_resvs,5,resvs_data[5])
                    resvs.write(r_resvs,6,resvs_data[6])
                    resvs.write(r_resvs,7,resvs_data[7].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")                      
            else:
                resvs.write(0,0,"No Entries in STOCK_RESVS")

            o_stock_updates.execute("""SELECT BU_CODE_LU,ITEM_NO,BU_CODE_SUP,ON_HAND_QTY,AVAIL_QTY,ALLOC_QTY,BLOCKED_QTY,RESERVED_QTY,EVENT_DESC,TRANS_ID,TRANS_STATUS,ERR_REASON,INS_DATE 
FROM O_STOCK_UPDATES 
WHERE ITEM_NO IN 
(SELECT DISTINCT ITEM_NO FROM I_STOCK_RESV_REQ_LINES WHERE TRANS_ID IN (SELECT TRANS_ID FROM I_STOCK_RESV_REQS WHERE ORD_ID_REF_SALES IN ("""+ord+"""))) AND 
BU_CODE_LU IN 
(SELECT DISTINCT BU_CODE_LU FROM I_STOCK_RESV_REQ_LINES WHERE TRANS_ID IN (SELECT TRANS_ID FROM I_STOCK_RESV_REQS  WHERE ORD_ID_REF_SALES IN ("""+ord+"""))) AND 
INS_DATE BETWEEN '"""+end_time+"""' AND '"""+start_time+"""' AND EVENT_DESC = 'STOCK RESERVATION' ORDER BY INS_DATE DESC""")
            o_stock_updates_data=o_stock_updates.fetchall()
            stk_upd = self.wb.add_worksheet("O_STOCK_UPDATES")
            upd=0
            if(o_stock_updates.rowcount>0):
                stk_upd.write(0,0,"BU_CODE_LU")
                stk_upd.write(0,1,"ITEM_NO")
                stk_upd.write(0,2,"BU_CODE_SUP")
                stk_upd.write(0,3,"ON_HAND_QTY")
                stk_upd.write(0,4,"AVAIL_QTY")
                stk_upd.write(0,5,"ALLOC_QTY")
                stk_upd.write(0,6,"BLOCKED_QTY")
                stk_upd.write(0,7,"RESERVED_QTY")
                stk_upd.write(0,8,"EVENT_DESC")
                stk_upd.write(0,9,"TRANS_ID")
                stk_upd.write(0,10,"TRANS_STATUS")
                stk_upd.write(0,11,"ERR_REASON")
                stk_upd.write(0,12,"INS_DATE")
                for upd_data in o_stock_updates_data:
                    upd+=1
                    stk_upd.write(upd,0,upd_data[0])
                    stk_upd.write(upd,1,upd_data[1])
                    stk_upd.write(upd,2,upd_data[2])
                    stk_upd.write(upd,3,upd_data[3])
                    stk_upd.write(upd,4,upd_data[4])
                    stk_upd.write(upd,5,upd_data[5])
                    stk_upd.write(upd,6,upd_data[6])
                    stk_upd.write(upd,7,upd_data[7])
                    stk_upd.write(upd,8,upd_data[8])
                    stk_upd.write(upd,9,upd_data[9])
                    stk_upd.write(upd,10,upd_data[10])
                    stk_upd.write(upd,11,upd_data[11])
                    stk_upd.write(upd,12,upd_data[12].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
            else:
                stk_upd.write(0,0,"No Entries in O_STOCK_UPDATES")

        i_stock_resv_reqs.close()
        i_stock_resv_req_lines.close()
        o_stock_updates.close()
        stock_resv_reqs.close()
        stock_resv_req_lines.close()
        stock_resvs.close()
        return 1

class StockAllocation:
    def __init__(self,wb,wb_r):
        self.wb=wb
        self.wb_r=wb_r
    def processStockAllocation(self):
        global var1,e1,e2
        start_time=''
        end_time=''
        #Reading orders from input.xlsx
        ord=''
        orders = self.wb_r.sheet_by_name("Stock allocation")
        sales_orders = []
        if(var1==0):
            for i in range(1,orders.nrows):
                sales_orders.append(orders.cell_value(i,0))
        else:
            sales_orders = e1.get().split(",")

        for index in range(len(sales_orders)):
                ord+="'"+sales_orders[index]+"'"
                if(index<len(sales_orders)-1):
                    ord+=","
        
        i_orders = iip_obj.cursor()
        i_order_lines=iip_obj.cursor()
        o_stock_updates=iip_obj.cursor()
        order_lines=wis_data.cursor()
        display("Fetching Stock Allocations")
        carry_on=False
        i_orders.execute("SELECT TRANS_ID,ORD_ID_TK,BU_CODE_LU,ORDER_SRC,WRK_ORD_REF,ORG_ORD_REF,TRANS_STATUS,INS_DATE,ERR_REASON,ORD_ACTION,ORD_ID_REF_SALES,PAYMENT_OPTION FROM I_ORDERS WHERE ORD_ACTION IN ('ALLOCATE','ALLOCATED') AND INSTR(ORD_ID_REF_SALES,"+ord+",1,1)>0 ORDER BY INS_DATE DESC")
        i_orders_data=i_orders.fetchall()
        i_ord = self.wb.add_worksheet("I_ORDERS")
        ord_r=0
        if(i_orders.rowcount>0):
            i_ord.write(0,0,"TRANS_ID")
            i_ord.write(0,1,"ORD_ID_TK")
            i_ord.write(0,2,"BU_CODE_LU")
            i_ord.write(0,3,"ORDER_SRC")
            i_ord.write(0,4,"WRK_ORD_REF")
            i_ord.write(0,5,"ORG_ORD_REF")
            i_ord.write(0,6,"TRANS_STATUS")
            i_ord.write(0,7,"INS_DATE")
            i_ord.write(0,8,"ERR_REASON")
            i_ord.write(0,9,"ORD_ACTION")
            i_ord.write(0,10,"ORD_ID_REF_SALES")
            i_ord.write(0,11,"PAYMENT_OPTION")
            for ord_data in  i_orders_data:
                ord_r+=1
                i_ord.write(ord_r,0,ord_data[0])
                i_ord.write(ord_r,1,ord_data[1])
                i_ord.write(ord_r,2,ord_data[2])
                i_ord.write(ord_r,3,ord_data[3])
                i_ord.write(ord_r,4,ord_data[4])
                i_ord.write(ord_r,5,ord_data[5])
                i_ord.write(ord_r,6,ord_data[6])
                i_ord.write(ord_r,7,ord_data[7].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                i_ord.write(ord_r,8,ord_data[8])
                i_ord.write(ord_r,9,ord_data[9])
                i_ord.write(ord_r,10,ord_data[10])
                i_ord.write(ord_r,11,ord_data[11])
                if(ord_data[6]==7):
                    carry_on=True
        else:
            i_ord.write(0,0,"No Entries in I_ORDERS")
        if(carry_on):
            i_order_lines.execute("SELECT ORD_ID_TK,ITEM_NO,ITEM_QTY,BU_CODE_SUP,INS_DATE,ORG_ORDL_REF FROM I_ORDER_LINES WHERE ORD_ID_TK IN (SELECT ORD_ID_TK FROM I_ORDERS WHERE ORD_ACTION IN ('ALLOCATE','ALLOCATED') AND INSTR(ORD_ID_REF_SALES,"+ord+",1,1)>0) ORDER BY INS_DATE DESC")
            i_order_lines_data = i_order_lines.fetchall()
            i_ordl=self.wb.add_worksheet("I_ORDER_LINES")
            ordl_r=0
            if(i_order_lines.rowcount>0):
                i_ordl.write(0,0,"ORD_ID_TK")
                i_ordl.write(0,1,"ITEM_NO")
                i_ordl.write(0,2,"ITEM_QTY")
                i_ordl.write(0,3,"BU_CODE_SUP")
                i_ordl.write(0,4,"INS_DATE")
                i_ordl.write(0,5,"ORG_ORDL_REF")
                for ordl_data in i_order_lines_data:
                    ordl_r+=1
                    i_ordl.write(ordl_r,0,ordl_data[0])
                    i_ordl.write(ordl_r,1,ordl_data[1])
                    i_ordl.write(ordl_r,2,ordl_data[2])
                    i_ordl.write(ordl_r,3,ordl_data[3])
                    i_ordl.write(ordl_r,4,ordl_data[4].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                    i_ordl.write(ordl_r,5,ordl_data[5])
        if(carry_on):
            carry_on=False
            order_lines.execute("SELECT ORD_ID,BU_TK_SUP,ITEM_NO,ITEM_QTY,ITEM_QTY_ORIG,ORDL_STATUS,INS_DATE,ORD_ID_IN_TK,SS_RUN_TK,LU_TK,BU_CODE_SUP,ORG_ORDL_REF,ALLOC_QTY FROM ORDER_LINES WHERE ORD_ID_IN_TK IN (SELECT ORD_ID_TK FROM I_ORDERS WHERE ORD_ACTION IN ('ALLOCATE','ALLOCATED') AND INSTR(ORD_ID_REF_SALES,"+ord+",1,1)>0) ORDER BY INS_DATE DESC")
            order_lines_data = order_lines.fetchall()
            ordl = self.wb.add_worksheet("ORDER_LINES")
            ordl_r=0
            if(order_lines.rowcount>0):
                carry_on=True
                ordl.write(0,0,"ORD_ID")
                ordl.write(0,1,"BU_TK_SUP")
                ordl.write(0,2,"ITEM_NO")
                ordl.write(0,3,"ITEM_QTY")
                ordl.write(0,4,"ITEM_QTY_ORIG")
                ordl.write(0,5,"ORDL_STATUS")
                ordl.write(0,6,"INS_DATE")
                ordl.write(0,7,"ORD_ID_IN_TK")
                ordl.write(0,8,"SS_RUN_TK")
                ordl.write(0,9,"LU_TK")
                ordl.write(0,10,"BU_CODE_SUP")
                ordl.write(0,11,"ORG_ORDL_REF")
                ordl.write(0,12,"ALLOC_QTY")
                for ordl_data in order_lines_data:
                    if(ordl_r==0):
                        end_time = (ordl_data[6]+timedelta(minutes=1)).strftime("%d-%b-%y %I.%M.%S.%f %p").upper()+" +00:00"
                    if(ordl_r==(order_lines.rowcount-1)):
                        start_time = (ordl_data[6]-timedelta(minutes=1)).strftime("%d-%b-%y %I.%M.%S.%f %p").upper()+" +00:00"
                    ordl_r+=1
                    ordl.write(ordl_r,0,ordl_data[0])
                    ordl.write(ordl_r,1,ordl_data[1])
                    ordl.write(ordl_r,2,ordl_data[2])
                    ordl.write(ordl_r,3,ordl_data[3])
                    ordl.write(ordl_r,4,ordl_data[4])
                    ordl.write(ordl_r,5,ordl_data[5])
                    ordl.write(ordl_r,6,ordl_data[6].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")
                    ordl.write(ordl_r,7,ordl_data[7])
                    ordl.write(ordl_r,8,ordl_data[8])
                    ordl.write(ordl_r,9,ordl_data[9])
                    ordl.write(ordl_r,10,ordl_data[10])
                    ordl.write(ordl_r,11,ordl_data[11])
                    ordl.write(ordl_r,12,ordl_data[12])
            else:
                ordl.write(0,0,"No Entries in ORDER_LINES")

        if(carry_on):
            o_stock_updates.execute("""SELECT BU_CODE_LU,ITEM_NO,BU_CODE_SUP,ORD_SRC,ON_HAND_QTY,AVAIL_QTY,ALLOC_QTY,BLOCKED_QTY,RESERVED_QTY,EVENT_DESC,TRANS_ID,TRANS_STATUS,ERR_REASON,INS_DATE FROM O_STOCK_UPDATES WHERE 
ITEM_NO IN (SELECT DISTINCT ITEM_NO FROM I_ORDER_LINES WHERE ORD_ID_TK IN (SELECT ORD_ID_TK FROM I_ORDERS WHERE ORD_ACTION IN ('ALLOCATE','ALLOCATED') AND INSTR(ORD_ID_REF_SALES,'330050408',1,1)>0)) AND
BU_CODE_LU IN (SELECT DISTINCT BU_CODE_LU FROM I_ORDERS WHERE ORD_ACTION IN ('ALLOCATE','ALLOCATED') AND INSTR(ORD_ID_REF_SALES,"""+ord+""",1,1)>0) AND
INS_DATE BETWEEN '"""+start_time+"""' AND '"""+end_time+"""' AND
EVENT_DESC = 'CREATE ORDER'""")
            o_stock_updates_data=o_stock_updates.fetchall()
            stk_upd=self.wb.add_worksheet("O_STOCK_UPDATES")
            upd=0
            if(o_stock_updates.rowcount>0):
                stk_upd.write(0,0,"BU_CODE_LU")
                stk_upd.write(0,1,"ITEM_NO")
                stk_upd.write(0,2,"BU_CODE_SUP")
                stk_upd.write(0,3,"ORD_SRC")
                stk_upd.write(0,4,"ON_HAND_QTY")
                stk_upd.write(0,5,"AVAIL_QTY")
                stk_upd.write(0,6,"ALLOC_QTY")
                stk_upd.write(0,7,"BLOCKED_QTY")
                stk_upd.write(0,8,"RESERVED_QTY")
                stk_upd.write(0,9,"EVENT_DESC")
                stk_upd.write(0,10,"TRANS_ID")
                stk_upd.write(0,11,"TRANS_STATUS")
                stk_upd.write(0,12,"ERR_REASON")
                stk_upd.write(0,13,"INS_DATE")
                for upd_data in o_stock_updates_data:
                    upd+=1
                    stk_upd.write(upd,0,upd_data[0])
                    stk_upd.write(upd,1,upd_data[1])
                    stk_upd.write(upd,2,upd_data[2])
                    stk_upd.write(upd,3,upd_data[3])
                    stk_upd.write(upd,4,upd_data[4])
                    stk_upd.write(upd,5,upd_data[5])
                    stk_upd.write(upd,6,upd_data[6])
                    stk_upd.write(upd,7,upd_data[7])
                    stk_upd.write(upd,8,upd_data[8])
                    stk_upd.write(upd,9,upd_data[9])
                    stk_upd.write(upd,10,upd_data[10])
                    stk_upd.write(upd,11,upd_data[11])
                    stk_upd.write(upd,12,upd_data[12])
                    stk_upd.write(upd,13,upd_data[13].strftime("%d-%b-%y %I.%M.%S.%f %p")+" +00:00")                 
            else:
                stk_upd.write(0,0,"No Entries in O_STOCK_UPDATES") 

        i_orders.close()
        i_order_lines.close()
        o_stock_updates.close()
        order_lines.close()
        
        return 1

class PushToQueue:
    def __init__(self,wb):
        self.wb=wb
    def push_xml(self):
        global e1,e2
        in_exe_data=''
        que_exe_data=''
        out_exe_data=''
        xml = open("xml_input.xml","r").read()
        push_to_out_queue=iip_obj.cursor()
        xml_message_log = iip_obj.cursor()
        out_queue = iip_obj.cursor()
        display("push xml to queue")
        proc = """
        DECLARE
            v_queue_name          iip_types.t_chr_value := '"""+e1.get()+"""';
            v_nr_msgs             number                := """+e2.get()+""";
            v_payload             iip_types.t_xmltype;
            v_queue_options       dbms_aq.ENQUEUE_OPTIONS_T;
            v_message_properties  dbms_aq.MESSAGE_PROPERTIES_T;
            v_message_id          iip_types.t_raw;
            i                     number;
        BEGIN
            dbms_output.put_line('Start time->'||systimestamp);

            v_payload := xmltype.createxml('"""+xml+"""');
            
            FOR i IN 1..v_nr_msgs LOOP
            dbms_aq.enqueue(queue_name => v_queue_name
                            ,enqueue_options => v_queue_options
                            ,message_properties => v_message_properties
                            ,payload => v_payload
                            ,msgid => v_message_id);
            END LOOP;
                    dbms_output.put_line('End time->'||systimestamp); 
        --      dbms_output.put_line('Message ID: '||v_message_id);                   
        --    COMMIT;
        EXCEPTION
            WHEN OTHERS THEN
        --      ROLLBACK;
            dbms_output.put_line(SQLERRM);
        END;"""
        push_to_out_queue.execute(proc)
        iip_obj.commit()
        time.sleep(10)
        push_to_out_queue.close()
        in_tree = ET.parse("xml_input.xml")
        in_root = in_tree.getroot()
        in_root_exp=re.compile(r'{.+}').findall(str(in_root))[0]
        if(e1.get()=='IIP_OBJ.Q_IFC_OUT_001_QUE'):
            in_exe = in_root.findall('.//'+in_root_exp+'ExecutableOrderId')[0]
            if(in_exe not in [None]):
                in_exe_data=in_exe.text
        elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_002_QUE'):
            in_exe=in_root.find('.//'+in_root_exp+'ManualOrderRequestId')
            if(in_exe not in [None]):
                in_exe_data=in_exe.text
        elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_004_QUE'):
            in_exe=in_root.findall('.//'+in_root_exp+'ExecutableOrderId')[0]
            if(in_exe not in [None]):
                in_exe_data=in_exe.text
        else:
            pass
        ws_out=self.wb.add_worksheet("TRIGGERED STATUS")
        carry_on=False
        if(e1.get()=='IIP_OBJ.Q_IFC_OUT_001_QUE'):
            out_queue.execute("SELECT q.USER_DATA.GETCLOBVAL(),q.ENQ_TIME FROM Q_IFC_OUT_001 q WHERE TRUNC(q.ENQ_TIME) LIKE SYSDATE ORDER BY q.ENQ_TIME DESC")
            out_queue_data=out_queue.fetchall()
        elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_002_QUE'):
            out_queue.execute("SELECT q.USER_DATA.GETCLOBVAL(),q.ENQ_TIME FROM Q_IFC_OUT_002 q WHERE TRUNC(q.ENQ_TIME) LIKE SYSDATE ORDER BY q.ENQ_TIME DESC")
            out_queue_data=out_queue.fetchall()
        elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_004_QUE'):
            out_queue.execute("SELECT q.USER_DATA.GETCLOBVAL(),q.ENQ_TIME FROM Q_IFC_OUT_004 q WHERE TRUNC(q.ENQ_TIME) LIKE SYSDATE ORDER BY q.ENQ_TIME DESC")
            out_queue_data=out_queue.fetchall()
        else:
            pass
        if(out_queue.rowcount>0):
            ws_out.write(0,0,"Status")
            ws_out.write(0,1,"Executable Order id")
            ws_out.write(0,2,"ENQ_TIME")
            for queue in out_queue_data:
                with open('xml_msg.xml', 'w') as f:
                    f.write(queue[0].read())
                enq_time=queue[1].strftime("%d-%b-%y %I.%M.%S.%f %p")
                que_tree = ET.parse("xml_msg.xml")
                que_root=que_tree.getroot()
                que_root_exp=re.compile(r'{.+}').findall(str(que_root))[0]
                if(e1.get()=='IIP_OBJ.Q_IFC_OUT_001_QUE'):
                    que_exe = que_root.findall('.//'+que_root_exp+'ExecutableOrderId')[0]
                    if(que_exe not in [None]):
                        que_exe_data = que_exe.text
                elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_002_QUE'):
                    que_exe=que_root.find('.//'+que_root_exp+'ManualOrderRequestId')
                    if(que_exe not in [None]):
                        que_exe_data = que_exe.text
                elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_004_QUE'):
                    que_exe=que_root.findall('.//'+que_root_exp+'ExecutableOrderId')[0]
                    if(que_exe not in [None]):
                        que_exe_data = que_exe.text
                else:
                    pass
                if(in_exe_data==que_exe_data):
                    carry_on=True
                    break

        if(carry_on==True):
            ws_out.write(1,0,"Order is stuck in out queue")
            ws_out.write(1,1,str(que_exe_data))
            ws_out.write(1,2,enq_time)
            carry_on=False
        else:
            carry_on=True
        
        if(e1.get()=='IIP_OBJ.Q_IFC_IN_001_QUE'):
                ws_out.write(0,0,"Order is inserted, please check in tables!..")
                carry_on=False
    
        if(carry_on==True):
            carry_on=False
            if(e1.get()=='IIP_OBJ.Q_IFC_OUT_001_QUE'):
                xml_message_log.execute("SELECT log.TRANS_ID,log.TRANS_STATUS,log.XML_MESSAGE.GETCLOBVAL(),log.ERROR_MESSAGE,log.INS_DATE FROM XML_MESSAGE_LOG log WHERE log.MESSAGE_TYPE='SyncExecutableOrderG2Msg' AND TRUNC(log.INS_DATE) LIKE SYSDATE ORDER BY log.INS_DATE DESC")
                xml_message_log_data = xml_message_log.fetchall()
            elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_002_QUE'):
                xml_message_log.execute("SELECT log.TRANS_ID,log.TRANS_STATUS,log.XML_MESSAGE.GETCLOBVAL(),log.ERROR_MESSAGE,log.INS_DATE FROM XML_MESSAGE_LOG log WHERE log.MESSAGE_TYPE='RemoveManualOrderMsg' AND TRUNC(log.INS_DATE) LIKE SYSDATE ORDER BY log.INS_DATE DESC")
                xml_message_log_data = xml_message_log.fetchall()
            elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_004_QUE'):
                xml_message_log.execute("SELECT log.TRANS_ID,log.TRANS_STATUS,log.XML_MESSAGE.GETCLOBVAL(),log.ERROR_MESSAGE,log.INS_DATE FROM XML_MESSAGE_LOG log WHERE log.MESSAGE_TYPE='ModifyExecutableOrderLineMsg' AND TRUNC(log.INS_DATE) LIKE SYSDATE ORDER BY log.INS_DATE DESC")
                xml_message_log_data = xml_message_log.fetchall()
            else:
                pass
            if(xml_message_log.rowcount>0):
                ws_out.write(0,0,"Status")
                ws_out.write(0,1,"Order Ref")
                ws_out.write(0,2,"TRANS_ID")
                ws_out.write(0,3,"TRANS_STATUS")
                ws_out.write(0,4,"ERROR_MESSAGE")
                ws_out.write(0,5,"INS_DATE")
                for data in xml_message_log_data:
                    with open('xml_msg.xml', 'w') as f: 
                        f.write(data[2].read())
                    trans_id=data[0]
                    trans_status = data[1]
                    err_msg = data[3]
                    ins_date = data[4].strftime("%d-%b-%y %I.%M.%S.%f %p")
                    tree = ET.parse("xml_msg.xml")
                    root=tree.getroot()
                    root_exp=re.compile(r'{.+}').findall(str(root))[0]
                    if(e1.get()=='IIP_OBJ.Q_IFC_OUT_001_QUE'):
                        out_exe = root.findall('.//'+root_exp+'ExecutableOrderId')[0]
                        if(out_exe not in [None]):
                            out_exe_data=out_exe.text
                    elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_002_QUE'):
                        out_exe=root.find('.//'+root_exp+'ManualOrderRequestId')
                        if(out_exe not in [None]):
                            out_exe_data=out_exe.text
                    elif(e1.get()=='IIP_OBJ.Q_IFC_OUT_004_QUE'):
                        out_exe=root.findall('.//'+root_exp+'ExecutableOrderId')[0]
                        if(out_exe not in [None]):
                            out_exe_data=out_exe.text
                    else:
                        pass
                    if(in_exe_data==out_exe_data):
                        carry_on=True
                        break
            if(carry_on==True):
                ws_out.write(1,0,"Record is sent")
                ws_out.write(1,1,out_exe_data)
                ws_out.write(1,2,str(trans_id))
                ws_out.write(1,3,str(trans_status))
                ws_out.write(1,4,str(err_msg))
                ws_out.write(1,5,ins_date)                     
            else:
                ws_out.write(0,0,"No Record found in XML_MESSAGE_LOG")
        return 1

def onrun():
    closeOutput()
    global status,wb_r,wb,req_type,var1,e1,e2,process
    deleteContent("log.txt")
    clearDisplay()
    process.config(bg="orange")
    try:
        readInput()
        writeOutput()
        getConnections()
        full_sync_isom = ['Publish stock to ISOM','Onhand Inventory ISOM']
        full_sync_oms_gemini= ['Publish stock to OMS/GEMINI','Onhand inventory OMS/GEMINI']
        if(req_type=='Order Processing'):
            op=Orderprocess(wb,wb_r)
            status = op.processOrders()
        elif(req_type=='Inventory Details'):
            inv_details = InventoryDetails(wb,wb_r)
            status = inv_details.processInventory()
        elif(req_type in full_sync_isom):
            if(req_type=='Publish stock to ISOM'):
                stock_report_isom = PublishStockISOM(wb,wb_r,False)
            else:
                stock_report_isom = PublishStockISOM(wb,wb_r,True)
            status = stock_report_isom.processStockToISOM()
        elif(req_type in full_sync_oms_gemini):
            if(req_type=='Publish stock to OMS/GEMINI'):
                stock_report_oms = PublishStock_OMS_GEMINI(wb,wb_r,False)
            else:
                stock_report_oms = PublishStock_OMS_GEMINI(wb,wb_r,True)
            status=stock_report_oms.processStockToOMS()
        elif(req_type=='Cancellation'):
            cancellation = Cancellation(wb,wb_r)
            status = cancellation.processCancellation()
        elif(req_type=='Stock Adjustment'):
            stock_adjustment = StockAdjustments(wb,wb_r)
            status = stock_adjustment.process_adjustments()
        elif(req_type=='Stock Reservation'):
            stock_reservation = StockReservation(wb,wb_r)
            status = stock_reservation.processStockReservation()
        elif(req_type=='Stock Allocation'):
            stock_allocation=StockAllocation(wb,wb_r)
            status = stock_allocation.processStockAllocation()
        elif(req_type=='Push XML To Queue'):
            push = PushToQueue(wb)
            status = push.push_xml()
        else:
            pass
        closeConnections()  
    except Exception as e:
        with open("log.txt", 'a') as out:
            out.write(str(e))
        process.config(bg="red")
        os.startfile("log.txt")
    wb.close()
    if(status==1):
        process.config(bg="light green")
        os.startfile('output.xlsx')
    else:
        process.config(bg="red")
        os.startfile("log.txt")
    status=0   

#Creating GUI
req_type_var = StringVar(screen)
req_type_dict = ['Order Processing','Inventory Details','Publish stock to ISOM','Publish stock to OMS/GEMINI','Onhand Inventory ISOM','Onhand inventory OMS/GEMINI','Cancellation','Stock Adjustment','Stock Reservation','Stock Allocation','Push XML To Queue']
req_type_var.set('Order Processing')
e1.set("order1,order2,...")
e2.set("")
req_type = req_type_var.get()
req_type_menu = OptionMenu(screen, req_type_var, *req_type_dict)
req_type_menu.grid(row = 0, column =0)
def change_req_type(*args):
    global req_type,text1,text2,e1,e2,process,var1
    req_type = req_type_var.get()
    process.config(bg="grey")
    if(req_type in ['Order Processing','Cancellation','Publish stock to ISOM','Publish stock to OMS/GEMINI','Stock Reservation','Stock Allocation']):
        if(var1==1):
            text1.config(state=NORMAL)
            text2.config(state=DISABLED)
        else:
            text1.config(state=DISABLED)
            text2.config(state=DISABLED)
    elif(req_type in ['Inventory Details','Onhand Inventory ISOM','Onhand inventory OMS/GEMINI','Stock Adjustment','Push XML To Queue']):
        if(var1==1):
            text1.config(state=NORMAL)
            text2.config(state=NORMAL)
        else:
            text1.config(state=DISABLED)
            text2.config(state=DISABLED)
    else:
        pass
    
    if(req_type in ['Order Processing','Cancellation','Stock Reservation','Stock Allocation']):
        e1.set("order1,order2,...")
        e2.set("")
    elif(req_type in ['Publish stock to ISOM','Publish stock to OMS/GEMINI']):
        e1.set("cdc1,cdc2,...")
        e2.set("")
    elif(req_type in ['Inventory Details','Onhand Inventory ISOM','Onhand inventory OMS/GEMINI','Stock Adjustment']):
        e1.set("cdc")
        e2.set("article1,article2,...")
    elif(req_type == 'Push XML To Queue'):
        e1.set("QUEUE NAME")
        e2.set("Number of msgs")
    else:
        pass
    

   
req_type_var.trace('w', change_req_type)
env_var = StringVar(screen)
env_dict = ['CF0548','CF0549','CF0613','CF0614','CF0837','CF1021','CF1055','CF1057','CF1133','CF1183','CM0310','CM0322','CM0327','PP0878','PP1069','PP1420','PP1445','PP1453','PP1CWISAP','PT1466','PT1467','PT1520']
env_var.set('CF0837')
env=env_var.get()
env_menu = OptionMenu(screen, env_var, *env_dict)
env_menu.grid(row = 0, column =1)

def change_env(*args):
    global env
    env=env_var.get()
env_var.trace('w', change_env)
text1.grid(row=2,column=0)
text2.grid(row=2,column=1)
day_var = StringVar(screen)
day_list = ['Today','Latest']
day_var.set('Today')
day = day_var.get()
day_menu = OptionMenu(screen,day_var, *day_list)
day_menu.grid(row=0,column =2)

def change_day(*args):
    global day
    day=day_var.get()
day_var.trace('w', change_day)

chk=IntVar(screen)
Checkbutton(screen,text="Text Box",variable=chk,width=5).grid(row=2,column=2)

def change_chk(*args):
    global var1,input_btn,req_type
    var1=chk.get()
    if(var1==1):
        input_btn.config(state=DISABLED)
        text1.config(state=NORMAL)
        if(req_type in ['Inventory Details','Onhand Inventory ISOM','Onhand inventory OMS/GEMINI','Stock Adjustment','Push XML To Queue']):
            text2.config(state=NORMAL)
        else:
            text2.config(state=DISABLED)
    else:
        input_btn.config(state=NORMAL)
        text1.config(state=DISABLED)
        text2.config(state=DISABLED)

chk.trace('w',change_chk)
input_btn=Button(screen,text="Input Data",width=35,command=lambda:on_input(),bg="light blue")
input_btn.grid(row=4,column=0,sticky="we")
def on_input():
    global req_type
    if(req_type=='Push XML To Queue'):
        os.startfile('xml_input.xml')
    else:
        os.startfile('input.xlsx')
process=Button(screen,text="Process Data",width=35,command=lambda:onrun(),bg="gray")
process.grid(row=4,column=1,columnspan=2,sticky="we")
screen.mainloop()


