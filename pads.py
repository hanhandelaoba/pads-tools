from win32com import client as pads_client
import argparse, re, sys
import mysql.connector 

db_access_config = {
    'host':'192.168.0.100',
    'user':'wangkai',
    'password':'xxxxxx',
    'database':'db_name',
}
about_info = 'tools for pads schematic, developed by Wang Kai, ver:0.2'
valid_attrs_name = ('PART NO', 'Description', 'Geometry.Height', 'Manufacture 1', 'Manufacture 1 P/N', 'Not_In_Bom', 'Value') 

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
    # part_type_with_part_info_list = []
    # part_type_with_part_no_attributes = {'part_type':'', 'PART_NO':'', 'Description':'', 'Geometry_Height':'', 'Manufacture':'', 'Manufacture_PN':'', 'Value':''}
    logic_prj = powerlogic_app.ActiveDocument
    prj_part_types = logic_prj.PartTypes
    print(f'prj part types:{prj_part_types.Count}')
    prj_part_types.Sort()
    total_part_types = prj_part_types.Count
    deal_counter = 0
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
                print(f'\nget attributes from database error, part type:{part_type_name}')
                continue
            part_type_with_part_no_attributes['PART_NO'] = part_no
            part_type_with_part_no_attributes['Description'] = part_attributes['Description']
            part_type_with_part_no_attributes['Manufacture'] = part_attributes['Manufacture']
            part_type_with_part_no_attributes['Manufacture_PN'] = part_attributes['Manufacture_PN']
            # part_type_with_part_info_list.append(part_type_with_part_no_attributes)
            comps_with_same_part_type = prj_part_types.Item(j).Components
            for comp_with_same_part_type in comps_with_same_part_type:
                comp_with_same_part_type.Attributes.Item("Description").Value = part_attributes['Description']
                comp_with_same_part_type.Attributes.Item("Manufacture 1").Value = part_attributes['Manufacture']
                comp_with_same_part_type.Attributes.Item("Manufacture 1 P/N").Value = part_attributes['Manufacture_PN']
        else:
            print(f'\nhave part with invalid part no while supplementing parts attributes, part no:{part_no}, part type:{part_type_name}')
        deal_counter = deal_counter + 1
        print(f'\rtotal part types:{total_part_types},deal counter:{deal_counter}', end='')
    # print(part_type_with_part_info_list)
    # print(f'part_type_list length:{len(part_type_list)}')
    # print(f'selected part type is:{prj_part_types.Item(1).Name}')
    powerlogic_app.UnlockServer
    db_cnx.close()
    print(f'\nerror parts info while access database{errorPartsInfo}')
    print('supplement parts attributes info complete')

def clean_parts_attributes(powerlogic_app):    
    powerlogic_app.LockServer
    logic_prj = powerlogic_app.ActiveDocument 
    
    # collect part types in this prj
    part_type_list = []
    prj_part_types = logic_prj.PartTypes
    prj_part_types.Sort()
    for j in range(1, prj_part_types.Count + 1):
        part_type_list.append(prj_part_types.Item(j).Name)
    print(f'all part types in this prj({len(part_type_list)}):\n{part_type_list}')
    prj_comps = logic_prj.Components

    total_prj_comps = prj_comps.Count
    print(f'total components:{total_prj_comps}')
    clean_counter = 0    
    for prj_comp in prj_comps:    
        # collect part types in this prj
        # comp_part_type = prj_comp.PartTypeObject
        # if comp_part_type.Name not in part_type_list:
        #     part_type_list.append(comp_part_type.Name)
        
        # checkt part attributes
        comp_attrs = prj_comp.Attributes
        # delete invalid attributes
        for comp_attr in comp_attrs:            
            if comp_attr.Name not in valid_attrs_name:
                comp_attrs.Delete(comp_attr.Name)
        # add valid attributes
        if comp_attrs('Description') == None:
            comp_attrs.Add('Description', '')
        if comp_attrs('Geometry.Height') == None:
            comp_attrs.Add('Geometry.Height', '')        
        if comp_attrs('Manufacture 1') == None:
            comp_attrs.Add('Manufacture 1', '')
        if comp_attrs('Manufacture 1 P/N') == None:
            comp_attrs.Add('Manufacture 1 P/N', '')
        if comp_attrs('Not_In_Bom') == None:
            comp_attrs.Add('Not_In_Bom', '')
        if comp_attrs('PART NO') == None:
            comp_attrs.Add('PART NO', '')       
        if comp_attrs('Value') == None:
            comp_attrs.Add('Value', '')       

        comp_gates = prj_comp.Gates
        for comp_gate in comp_gates:
            comp_gate.SetVisibility(0, 'Value', 1)
            comp_gate.SetVisibility(3, '', 0)

        clean_counter = clean_counter + 1
        print(f'\rclean counter:{clean_counter}', end='')
        
    powerlogic_app.UnlockServer
    
    print('\nclean part attributes complete')

def check_parts_attributes(powerlogic_app):
    powerlogic_app.LockServer
    logic_prj = powerlogic_app.ActiveDocument 
    prj_comps = logic_prj.Components
    comps_with_invalid_part_no = []
    part_type_with_invalid_part_no = []
    total_prj_comps = prj_comps.Count
    print(f'total components:{total_prj_comps}')
    check_counter = 0 
    for prj_comp in prj_comps:            
        # check part attributes
        comp_attrs = prj_comp.Attributes
        comp_part_no = comp_attrs.Item("PART NO").Value  
        comp_part_no_length = len(comp_part_no)        
        comp_attr_no_in_bom = comp_attrs.Item("Not_In_Bom").Value     
        if comp_part_no != "NC" and comp_attr_no_in_bom != 'X' and comp_part_no_length != 8:
            # print(f'invalid PART NO value:{prj_comp.Name}')            
            comps_with_invalid_part_no.append(prj_comp.Name)
            comp_part_type = prj_comp.PartTypeObject
            if comp_part_type.Name not in part_type_with_invalid_part_no:
                part_type_with_invalid_part_no.append(comp_part_type.Name)
        check_counter = check_counter + 1
        print(f'\rcheck counter:{check_counter}', end='')

    powerlogic_app.UnlockServer
    print(f'\ncomponents with invalid part no({len(comps_with_invalid_part_no)}):\n{comps_with_invalid_part_no}\n')
    part_type_with_invalid_part_no_count = len(part_type_with_invalid_part_no)
    print(f'part type with invalid part no({part_type_with_invalid_part_no_count}):\n{part_type_with_invalid_part_no}')
    print('check part attributes complete')
    return part_type_with_invalid_part_no_count

def output_partlist():
    pass

if __name__ == '__main__':
    args_parser = argparse.ArgumentParser(prog='pads tools', description='check Schematic and Part Attributes, output partlist to excel for check')
    args_parser.add_argument('projectName', help='project name to deal with')
    args_parser.add_argument('--about', help=about_info, action='store_true')
    args_parser.add_argument('-a','--attr', help='clean parts attributes of opening schematic(将不使用的属性字段从原理图中删除)', action='store_true')
    args_parser.add_argument('-c','--check', help='check parts attributes information of opening schematic(检查原理图中所有器件是否分配了PART NO)', action='store_true')
    args_parser.add_argument('-s','--supplement', help='supplement parts attributes information of opening schematic(根据提供的PART NO信息，从数据库检索该器件的属性信息，并添加到每个器件里)', action='store_true')
    args_parser.add_argument('-p','--partlist', help='output parts list to excel', action='store_true')
    user_args = args_parser.parse_args()
        
    logic_app = pads_client.Dispatch('PowerLogic.Application')
    
    print(logic_app.Name)
    logic_prj = logic_app.ActiveDocument 
    print(logic_prj.Name)
    pattern = re.compile(user_args.projectName, re.I)
    match_result = pattern.match(logic_prj.Name)
    if match_result is None:        
        print(print('project name invalid'))
        sys.exit(0)
    
    print(f'project name valid:{match_result.group()}')  

    if user_args.attr:
        print('clean parts attributes of opening schematic')
        clean_parts_attributes(logic_app)
        sys.exit(0)
    if user_args.check:
        print('check parts attributes of opening schematic')
        check_parts_attributes(logic_app)
        sys.exit(0)
    if user_args.supplement:
        print('supplement parts attributes info')
        supplement_parts_attributes_info(logic_app)
        sys.exit(0) 
    if user_args.partlist:
        print('output parts list to excel')
        sys.exit(0)        
    if user_args.about:
        print(about_info)
        sys.exit(0)  