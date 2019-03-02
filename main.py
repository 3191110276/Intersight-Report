from __future__ import print_function
import time
import requests
import json
import intersight
from intersight.intersight_api_client import IntersightApiClient
from openpyxl import Workbook

# Creates Intersight API instance
api_instance = IntersightApiClient(
    private_key="./keyfile.txt",
    api_key_id=open('./apikey.txt',"r").read(),
)

# Initiates the Intersight API connectors
ComputeRackUnitApi = intersight.ComputeRackUnitApi(api_instance)
EquipmentPsuApi = intersight.EquipmentPsuApi(api_instance)
BiosUnitApi	= intersight.BiosUnitApi(api_instance)
FirmwareRunningFirmwareApi = intersight.FirmwareRunningFirmwareApi(api_instance)
ManagementControllerApi = intersight.ManagementControllerApi(api_instance)
ComputeBoardApi = intersight.ComputeBoardApi(api_instance)
ProcessorUnitApi = intersight.ProcessorUnitApi(api_instance)
MemoryArrayApi = intersight.MemoryArrayApi(api_instance)
    
wb = Workbook()

# Creates the worksheets
ws_server = wb.active
ws_server.title = "Server"
ws_memory = wb.create_sheet("Memory")
ws_networkadapter = wb.create_sheet("Network Adapter")
ws_pcie = wb.create_sheet("PCIe Devices")
ws_disk = wb.create_sheet("Disks")
ws_psu = wb.create_sheet("PSUs")

# Creates table headers
ws_server.append(['Server ID', 'User Label', 'Asset Tag', 'Serial', 'DN', 'Model', 'Oper Power State', 'MOID', 'UUID', 'Platform Type', 'Service Profile', 'KVM IP', 'Num of CPUs', 'CPU Model', 'CPU Architecture', 'Num of CPU Cores', 'Num of CPU Cores enabled', 'Num of Threads', 'CPU Speed', 'CPU Stepping', 'Total Memory', 'Memory Speed', 'Num of Adaptors', 'Num of Eth Host Int', 'Num of FC Host Int', 'BIOS Firmware Version', 'BMC Firmware System', 'BMC Firmware Bootloader'])
ws_memory.append([])
ws_networkadapter.append([])
ws_pcie.append([])
ws_disk.append([])
ws_psu.append(['Serial', 'DN', 'Server MOID (Parent)', 'Model', 'MOID', 'PSU ID', 'PSU Presence'])

racks = ComputeRackUnitApi.compute_rack_units_get().results

for r in racks:
    bios_fw = FirmwareRunningFirmwareApi.firmware_running_firmwares_moid_get(BiosUnitApi.bios_units_moid_get(r.biosunits[0].moid).running_firmware[0].moid)
    bmc = ManagementControllerApi.management_controllers_moid_get(r.bmc.moid).running_firmware
    
    bmc_fw = {
        'system': None,
        'boot-loader': None
    }
    for b in bmc:
        x = FirmwareRunningFirmwareApi.firmware_running_firmwares_moid_get(b.moid)
        if(x.component == 'boot-loader'):
            bmc_fw['boot-loader'] = x.version
        elif(x.component == 'system'):
            bmc_fw['system'] = x.version
            
    compute_board = ComputeBoardApi.compute_boards_moid_get(r.board.moid)
    p = ProcessorUnitApi.processor_units_moid_get(compute_board.processors[0].moid)
    
    for m in compute_board.memory_arrays:
        print(MemoryArrayApi.memory_arrays_moid_get(m.moid))
        
    print(compute_board.presence)
    print(compute_board.serial)
    print(compute_board.storage_controllers)
    print(compute_board.storage_flex_flash_controllers)
    print(compute_board.storage_flex_util_controllers)
    
    ws_server.append([r.server_id, r.user_label, r.asset_tag, r.serial, r.dn, r.model, r.oper_power_state, r.moid, r.uuid, r.platform_type, r.service_profile, r.kvm_ip_addresses[0].address, r.num_cpus, p.model, p.architecture, r.num_cpu_cores, r.num_cpu_cores_enabled, r.num_threads, p.speed, p.stepping, r.total_memory, r.memory_speed, r.num_adaptors, r.num_eth_host_interfaces, r.num_fc_host_interfaces, bios_fw.version, bmc_fw['system'], bmc_fw['boot-loader']])
    
    #TODO: CPU model in this list
    
    #print(r.mod_time)
    #print(r.board)
    #print(r.tags)
    #print(r.locator_led)
    #print(r.registered_device)
    #print(r.rack_enclosure_slot)
    #print(r.storage_enclosures)
    #print(r.object_type)
    #print(r.fanmodules)
    #print(bios_fw.package_version)
    #print(bios_fw.registered_device)
    #print(r.device_mo_id)
    #print(compute_board.graphics_cards)
    #print(compute_board.equipment_tpms)
    
    for p in r.psus:
        psu = EquipmentPsuApi.equipment_psus_moid_get(p.moid)
        ws_psu.append([psu.serial, psu.dn, r.moid, psu.model, psu.moid, psu.psu_id, psu.presence])
        
        #print(psu.equipment_chassis)
        #print(psu.equipment_rack_enclosure)
        #print(psu.mod_time)
        #print(psu.device_mo_id)
    
    print('--------------')
    print(r.adapters)
    print('--------------')
    print(r.pci_devices)
    print('--------------')
    print(r.sas_expanders)
    print('--------------')
    print(r.generic_inventory_holders)
    
    
# Save the Excel file
wb.save("intersight_report.xlsx")