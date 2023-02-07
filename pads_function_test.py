from win32com import client as pads_client
import argparse
import re, sys

part_index  = 51

part_type_list = []
parts_list_with_same_part_type = []
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='check Schematic and Part Attributes, output partlist to excel for check')
    parser.add_argument('projectName', help='project name to deal with')
    parser.add_argument('-c','--check', help='check parts attributes of opening schematic', action='store_true')
    parser.add_argument('-p','--partlist', help='output parts list to excel', action='store_true')
    args = parser.parse_args()
    args_dict = vars(args)
    print(args_dict)
    if args.check:
        print('checkt schematic')
    if args.partlist:
        print('output parts list to excel')    

    logic_app = pads_client.Dispatch('PowerLogic.Application')    
    # logic_app = pads_client.Dispatch('PowerPCB.Application')
    print(logic_app.Name)
    logic_prj = logic_app.ActiveDocument 
    print(logic_prj.Name)
    pattern = re.compile(args.projectName, re.I)
    match_result = pattern.match(logic_prj.Name)
    if match_result is None:        
        print(print('project name invalid'))
        sys.exit(0)
    
    print(f'project name valid:{match_result.group()}')  
        
    logic_app.LockServer
    prj_comps = logic_prj.Components
    # comps_part_types = logic_prj.AttributeTypes
    # print(comps_part_types.Count)
    # for i in range(1, comps_part_types.Count +1):
    #     print(comps_part_types.Item(i))
    comp_part_type = prj_comps(part_index).PartTypeObject
    comp_part_name = prj_comps(part_index).Name
    comp_gates = prj_comps(part_index).Gates
    # print(comp_part_name)
    print(comp_part_type.Name)
    comps_this_part_type = comp_part_type.Components
    print(f'comps_this_part_type:{comps_this_part_type.Count}')
    for i in range(1, comps_this_part_type.Count + 1):
        # print(comps_this_part_type.Item(i).Name)
        parts_list_with_same_part_type.append(comps_this_part_type.Item(i).Name)
    # print(comp_gates.Count)
    # print(comp_gates.Item(1).GetVisibility(0, 'Value'))
    # for comp_gate in comp_gates:
    #     comp_gate.SetVisibility(3, '', 1)        
    #     print(comp_gate.GetVisibility(0, 'PART NO'))
    print(parts_list_with_same_part_type)
    logic_app.UnlockServer
    print('complete')

    # part_no = prj_comp.Attributes.Item("PART NO")
        # # print(part_no)
        # part_no_str = str(part_no)
        # if part_no_str == '10025376':
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
#     if pcb_comp.Attributes.Item("PART NO") == '10024519':
#         print(pcb_comp.name)
#         print(pcb_comp.Attributes.Item("PART NO"))
#         print(pcb_comp.Attributes.Item("Value"))    
#         print(pcb_comp.Attributes.Item("Manufacture 1")) 
#         print(pcb_comp.Attributes.Item("Manufacture 1 P/N"))
#         print(pcb_comp.Attributes.Item("Geometry.Height"))
#         print(pcb_comp.Attributes.Item("Description"))