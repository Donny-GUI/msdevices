# msdevices
Microsoft Device Manager module for python. Access all device names and properties with classes. 
Properly typed and hinted. 

# Classes 

## device 
string subclass for hinting

## name / device.name
device property name,  string subclass

## property / device.property
string of the device property's value


# Getting Started

```Bash

git clone https://github.com/Donny-GUI/msdevices.git

```

```Python3
from msdevices.devices import Devices

devices = Devices()

print(devices.names)
print(devices.properties)
print(devices.values)
print(devices.device_ids)
print(devices.array)

bluetooth_network_device = devices.device("Bluetooth Network Device")
if bluetooth_network_device:
  print(bluetooth_netork_decice.name)
  print(bluetooth_netork_decice.status_info)
  print(bluetooth_netork_decice.system_creation_class_name)
  print(bluetooth_netork_decice.availability)
  print(bluetooth_netork_decice.caption)
  print(bluetooth_netork_decice.class_guid)
  print(bluetooth_netork_decice.compatible_id)
  print(bluetooth_netork_decice.config_manager_error_code)
  print(bluetooth_netork_decice.device_id)
  # ... for all devices in __main__.devices.DevicePropertyNames

```
