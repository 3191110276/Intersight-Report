from __future__ import print_function
import time
import requests
import json
import intersight
from intersight.intersight_api_client import IntersightApiClient
from openpyxl import Workbook

api_instance = IntersightApiClient(
    private_key="./keyfile.txt",
    api_key_id=open('./apikey.txt',"r").read(),
)

rackApi = intersight.ComputeRackUnitApi(api_instance)

    
wb = Workbook()
ws = wb.active

# Creates table headers
ws.append(['Server ID', 'User Label', 'Model', 'Serial', 'DN', 'MOID', 'UUID'])

racks = rackApi.compute_rack_units_get().results

for r in racks:
    ws.append([r.server_id, r.user_label, r.model, r.serial, r.dn, r.moid, r.uuid])
    print(r.server_id)
    print(r.user_label)
    print(r.model)
    print(r.serial)
    print(r.dn)
    print(r.moid)
    print(r.uuid)


    
    print(r.adapters)
    print(r.asset_tag)
    print(r.available_memory)
    print(r.biosunits)
    print(r.bmc)
    print(r.board)
    print(r.device_mo_id)
    print(r.fanmodules)
    print(r.generic_inventory_holders)
    print(r.kvm_ip_addresses)
    print(r.locator_led)
    print(r.memory_speed)
    print(r.mod_time)
    print(r.num_adaptors)
    print(r.num_cpu_cores)
    print(r.num_cpu_cores_enabled)
    print(r.num_cpus)
    print(r.num_eth_host_interfaces)
    print(r.num_fc_host_interfaces)
    print(r.num_threads)
    print(r.object_type)
    print(r.oper_power_state)
    print(r.oper_state)
    print(r.operability)
    print(r.pci_devices)
    print(r.platform_type)
    print(r.psus)
    print(r.rack_enclosure_slot)
    print(r.registered_device)
    print(r.revision)
    print(r.rn)
    print(r.sas_expanders)
    print(r.service_profile)
    print(r.storage_enclosures)
    print(r.tags)
    print(r.total_memory)
    
# Save the Excel file
wb.save("intersight_report.xlsx")