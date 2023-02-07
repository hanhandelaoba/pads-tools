from win32com import client as pads_client
import argparse, re, sys

about_info = 'tools for pads schematic, developed by Wang Kai, ver:0.1'
valid_attrs_name = ('PART NO', 'Description', 'Geometry.Height', 'Manufacture 1', 'Manufacture 1 P/N', 'Not_In_Bom', 'Value') 

def clean_parts_attributes(powerlogic_app):    
    powerlogic_app.LockServer
    logic_prj = powerlogic_app.ActiveDocument 
    prj_comps = logic_prj.Components
    part_type_list = []
    for prj_comp in prj_comps:    
        # collect part types in this prj
        comp_part_type = prj_comp.PartTypeObject
        if comp_part_type.Name not in part_type_list:
            part_type_list.append(comp_part_type.Name)
        
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
    powerlogic_app.UnlockServer
    print(f'all part types in this prj({len(part_type_list)}):\n{part_type_list}')
    print('clean part attributes complete')

def check_parts_attributes(powerlogic_app):
    powerlogic_app.LockServer
    logic_prj = powerlogic_app.ActiveDocument 
    prj_comps = logic_prj.Components
    comps_with_invalid_part_no = []
    part_type_with_invalid_part_no = []
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
    powerlogic_app.UnlockServer
    print(f'\ncomponents with invalid part no({len(comps_with_invalid_part_no)}):\n{comps_with_invalid_part_no}\n')
    print(f'part type with invalid part no({len(part_type_with_invalid_part_no)}):\n{part_type_with_invalid_part_no}')
    print('check part attributes complete')

def output_partlist():
    pass

if __name__ == '__main__':
    args_parser = argparse.ArgumentParser(prog='pads tools', description='check Schematic and Part Attributes, output partlist to excel for check')
    args_parser.add_argument('projectName', help='project name to deal with')
    args_parser.add_argument('--about', help=about_info, action='store_true')
    args_parser.add_argument('-a','--attr', help='clean parts attributes of opening schematic', action='store_true')
    args_parser.add_argument('-c','--check', help='check parts attributes of opening schematic', action='store_true')
    args_parser.add_argument('-p','--partlist', help='output parts list to excel', action='store_true')
    user_args = args_parser.parse_args()
    # args_dict = vars(user_args)
    # print(args_dict)
    
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
    if user_args.partlist:
        print('output parts list to excel')
        sys.exit(0)        
    if user_args.about:
        print(about_info)
        sys.exit(0)  