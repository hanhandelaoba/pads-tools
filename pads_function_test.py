from win32com import client as pads_client
import argparse
import re, sys
import mysql.connector 

part_index  = 51
db_access_config = {
    'host':'192.168.0.100',
    'user':'wangkai',
    'password':'xxxxxx',
    'database':'db_name',
}
errorPartsInfo = []
def get_part_attributes_from_database(db_cnx, part_no):
    part_with_part_no_attributes = {
                    'PART_NO': '',
                    'Description': '',
                    'Manufacture_PN': '',
                    'Manufacture': '',
                    }
    
    try:
        db_cursorA = db_cnx.cursor(buffered=True)
        db_cursorB = db_cnx.cursor(buffered=True)

        # for part in bomPartsList:                        
        strQuery = 'SELECT part_desc, manufacture, order_partNo FROM wuliao_info WHERE sap_part_no=' + "'" + part_no + "'" + 'LIMIT 1'
        query = (strQuery)
        db_cursorA.execute(query,)
        partCount = db_cursorA.rowcount               
        db_cursor = db_cursorA
        if partCount == 0:
            strQuery = 'SELECT sap_desc, manufacture, mpn_desc FROM sap_mpn_map WHERE sap_part_no=' + "'" + part_no + "'" + 'LIMIT 1'
            query = (strQuery)
            db_cursorB.execute(query,)
            partCount = db_cursorA.rowcount
            db_cursor = db_cursorB
            partCountB = db_cursorB.rowcount                    
            if partCountB == 0:
                errorPartsInfo.append(part_no)    
                return None       
        for item in db_cursor:
            part_with_part_no_attributes = {
                    'PART_NO': part_no,
                    'Description': item[0],
                    'Manufacture': item[1],
                    'Manufacture_PN': item[2],                    
                    }     
            break          

    except Exception as err_db_access:            
            print(f'get_part_attributes_from_database error{str(err_db_access)}')

    db_cursorA.close()
    db_cursorB.close()
    
    return part_with_part_no_attributes

def supplement_parts_attributes_info(powerlogic_app):
    try:
        db_cnx = mysql.connector.connect(**db_access_config)
    except Exception as err:
        print(f'get_part_attributes_from_database error{str(err)}')
    powerlogic_app.LockServer
    part_type_with_part_info_list = []
    # part_type_with_part_no_attributes = {'part_type':'', 'PART_NO':'', 'Description':'', 'Geometry_Height':'', 'Manufacture':'', 'Manufacture_PN':'', 'Value':''}
    logic_prj = powerlogic_app.ActiveDocument
    prj_part_types = logic_prj.PartTypes
    print(f'prj part types:{prj_part_types.Count}')
    prj_part_types.Sort()
    for j in range(1, prj_part_types.Count + 1):
        part_type_with_part_no_attributes = {'part_type':'', 'PART_NO':'', 'Description':'', 'Geometry_Height':'', 'Manufacture':'', 'Manufacture_PN':'', 'Value':''}
        part_type_name = prj_part_types.Item(j).Name
        part_type_with_part_no_attributes['part_type'] = part_type_name        
        comp = prj_part_types.Item(j).Components.Item(1)
        # print(f'comp name:{comp.Name}')
        # print(comp.Attributes.Item("PART NO").Value)
        part_no = comp.Attributes.Item("PART NO").Value
        comp_part_no_length = len(part_no)
        # comp_attr_no_in_bom = comp.Attributes.Item("Not_In_Bom").Value
        if part_no != "NC" and comp_part_no_length == 8:        
            part_attributes = get_part_attributes_from_database(db_cnx=db_cnx, part_no=part_no)
            if part_attributes is None:
                print(f'get attributes from database error, part type:{part_type_name}')
                continue
            part_type_with_part_no_attributes['PART_NO'] = part_no
            part_type_with_part_no_attributes['Description'] = part_attributes['Description']
            part_type_with_part_no_attributes['Manufacture'] = part_attributes['Manufacture']
            part_type_with_part_no_attributes['Manufacture_PN'] = part_attributes['Manufacture_PN']
            part_type_with_part_info_list.append(part_type_with_part_no_attributes)
            comps_with_same_part_type = prj_part_types.Item(j).Components
            for comp_with_same_part_type in comps_with_same_part_type:
                comp_with_same_part_type.Attributes.Item("Description").Value = part_attributes['Description']
                comp_with_same_part_type.Attributes.Item("Manufacture 1").Value = part_attributes['Manufacture']
                comp_with_same_part_type.Attributes.Item("Manufacture 1 P/N").Value = part_attributes['Manufacture_PN']
        else:
            print(f'have part with invalid part no while supplementing parts attributes, part no:{part_no}, part type:{part_type_name}')
    print(part_type_with_part_info_list)
    # print(f'part_type_list length:{len(part_type_list)}')
    # print(f'selected part type is:{prj_part_types.Item(1).Name}')
    powerlogic_app.UnlockServer
    db_cnx.close()
    print(f'error parts info while access database{errorPartsInfo}')
    print('supplement_parts_attributes_info complete')

# parts_list_with_same_part_type = []
if __name__ == '__main__':
    args_parser = argparse.ArgumentParser(description='check Schematic and Part Attributes, output partlist to excel for check')
    args_parser.add_argument('projectName', help='project name to deal with')
    args_parser.add_argument('-c','--check', help='check parts attributes of opening schematic', action='store_true')
    args_parser.add_argument('-s','--supplement', help='supplement parts attributes information of opening schematic', action='store_true')
    args_parser.add_argument('-p','--partlist', help='output parts list to excel', action='store_true')
    args = args_parser.parse_args()
        
    logic_app = pads_client.Dispatch('PowerLogic.Application')    
    print(logic_app.Name)
    logic_prj = logic_app.ActiveDocument 
    print(logic_prj.Name)
    pattern = re.compile(args.projectName, re.I)
    match_result = pattern.match(logic_prj.Name)
    if match_result is None:        
        print(print('project name invalid'))
        sys.exit(0)
    
    print(f'project name valid:{match_result.group()}')    
    if args.check:
        print('checkt schematic')
    if args.supplement:
        print('supplement_parts_attributes_info')
        # sap = '10000120'
        # strQuery = 'SELECT part_desc, manufacture, order_partNo FROM wuliao_info WHERE sap_part_no=' + "'" + sap + "'"
        # print(strQuery)
        supplement_parts_attributes_info(logic_app)
    # comps_part_types = logic_prj.AttributeTypes
    # print(comps_part_types.Count)
    # for i in range(1, comps_part_types.Count +1):
    #     print(comps_part_types.Item(i))
    # comp_part_type = prj_comps(part_index).PartTypeObject
    # comp_part_name = prj_comps(part_index).Name
    # comp_gates = prj_comps(part_index).Gates
    # print(comp_part_name)
    # print(comp_part_type.Name)
    # comps_this_part_type = comp_part_type.Components
    # print(f'comps_this_part_type:{comps_this_part_type.Count}')
    # for i in range(1, comps_this_part_type.Count + 1):
        # print(comps_this_part_type.Item(i).Name)
        # parts_list_with_same_part_type.append(comps_this_part_type.Item(i).Name)
    # print(comp_gates.Count)
    # print(comp_gates.Item(1).GetVisibility(0, 'Value'))
    # for comp_gate in comp_gates:
    #     comp_gate.SetVisibility(3, '', 1)        
    #     print(comp_gate.GetVisibility(0, 'PART NO'))
    # print(parts_list_with_same_part_type)
    
    print('complete')

    # part_no = prj_comp.Attributes.Item("PART NO")
        # # print(part_no)
        # part_no_str = str(part_no)
        # if part_no_str == '11125376':
        #     print(prj_comp.name)
        #     print(prj_comp.Attributes.Item("PART NO"))
        #     print(prj_comp.Attributes.Item("Value"))    
        #     print(prj_comp.Attributes.Item("Manufacture 1")) 
        #     print(prj_comp.Attributes.Item("Manufacture 1 P/N"))
        #     print(prj_comp.Attributes.Item("Geometry.Height"))
        #     print(prj_comp.Attributes.Item("Description"))
        #     prj_comp.Attributes.Add('Voltage')
        #     prj_comp.Attributes.Add('Value', 'PM8937')
    # logic_app.Quit()

    # pcb_app = pads_client.Dispatch('powerpcb.application')
# print(pcb_app.name)
# pcb_prj = pcb_app.activedocument 
# print(pcb_prj.name)
# pcp_comps = pcb_prj.components
# for pcb_comp in pcp_comps:    
#     if pcb_comp.Attributes.Item("PART NO") == '11124519':
#         print(pcb_comp.name)
#         print(pcb_comp.Attributes.Item("PART NO"))
#         print(pcb_comp.Attributes.Item("Value"))    
#         print(pcb_comp.Attributes.Item("Manufacture 1")) 
#         print(pcb_comp.Attributes.Item("Manufacture 1 P/N"))
#         print(pcb_comp.Attributes.Item("Geometry.Height"))
#         print(pcb_comp.Attributes.Item("Description"))