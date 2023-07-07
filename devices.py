from win32com import client as windows_client


WMI_OBJECT: str = "winmgmts:"
PNP_ENTITY: str = "Win32_PnpEntity"


def get_devices():
    wmi = windows_client.GetObject(WMI_OBJECT)
    _devices = wmi.InstancesOf(PNP_ENTITY)
    devs = []
    # find bluetooth devices
    for _device in _devices:
        properties = [(prop.Name, prop.Value) for prop in _device.Properties_]
        props = {}

        for prop in properties:
            name, value = prop
            props[name] = value

        devs.append(Device(device.Name, props))
    return devs


DevicePropertyNames = [
    "StatusInfo", "SystemCreationClassName", "SystemName", "Availability", "Caption", "ClassGuid", "CompatibleID",
    "ConfigManagerErrorCode", "ConfigManagerUserConfig", "CreationClassName", "Description", "DeviceID", "ErrorCleared",
    "ErrorDescription", "HardwareID", "InstallDate", "LastErrorCode", "Manufacturer", "Name", "PNPClass", "PNPDeviceID",
    "PowerManagementCapabilities", "PowerManagementSupport", "Present", "Service", "Status"
]


class _name(property):
    pass


class _property(property):
    pass


class device:
    name = _name
    property = _property


class PropertyName(str):
    pass


class DeviceName(property):
    pass


class StatusInfo(property):
    pass


class SystemCreationClassName(property):
    pass


class SystemName(property):
    pass


class Availability(property):
    pass


class Caption(property):
    pass


class ClassGuid(property):
    pass


class CompatibleID(property):
    pass


class ConfigManagerErrorCode(property):
    pass


class ConfigManagerUserConfig(property):
    pass


class CreationClassName(property):
    pass


class Description(property):
    pass


class DeviceID(property):
    pass


class ErrorCleared(property):
    pass


class ErrorDescription(property):
    pass


class HardwareID(property):
    pass


class InstallDate(property):
    pass


class LastErrorCode(property):
    pass


class Manufacturer(property):
    pass


class Name(property):
    pass


class PNPClass(property):
    pass


class PNPDeviceID(property):
    pass


class PowerManagementCapabilities(property):
    pass


class PowerManagementSupport(property):
    pass


class Present(property):
    pass


class Service(property):
    pass


class Status(property):
    pass


class PropertyValue(str):
    pass


class Device(object):
    def __init__(self, name: str = None, properties: dict = None) -> None:
        self._name = name
        self._status_info: StatusInfo = properties["StatusInfo"]
        self._system_creation_class_name: SystemCreationClassName = properties["SystemCreationClassName"]
        self._system_name: SystemName = properties["SystemName"]
        self._availability: Availability = properties["Availability"]
        self._caption: Caption = properties["Caption"]
        self._class_guid: ClassGuid = properties["ClassGuid"]
        self._compatible_id: CompatibleID = properties["CompatibleID"]
        self._config_manager_error_code: ConfigManagerErrorCode = properties["ConfigManagerErrorCode"]
        self._config_manager_user_config: ConfigManagerUserConfig = properties["ConfigManagerUserConfig"]
        self._creation_class_name: CreationClassName = properties["CreationClassName"]
        self._description: Description = properties["Description"]
        self._device_id: DeviceID = properties["DeviceID"]
        self._error_cleared: ErrorCleared = properties["ErrorCleared"]
        self._error_description: ErrorDescription = properties["ErrorDescription"]
        self._hardware_id: HardwareID = properties["HardwareID"]
        self._install_date: InstallDate = properties["InstallDate"]
        self._last_error_code: LastErrorCode = properties["LastErrorCode"]
        self._manufacturer: Manufacturer = properties["Manufacturer"]
        self._name: Name = properties["Name"]
        self._pnp_class: PNPClass = properties["PNPClass"]
        self._pnp_device_id: PNPDeviceID = properties["PNPDeviceID"]
        self._power_management_capabilities: PowerManagementCapabilities = properties["PowerManagementCapabilities"]
        self._power_management_support: PowerManagementSupport = properties["PowerManagementSupport"]
        self._present: Present = properties["Present"]
        self._service: Service = properties["Service"]
        self._status: Status = properties["Status"]
        self._properties = properties

    def __str__(self):
        print(self.name)

    @property
    def names(self) -> list[device.name]:
        """ Returns the names of the properties """
        return list(self._properties.keys())

    @property
    def keys(self) -> list[device.name]:
        """ Returns the keys of the properties"""
        return list(self._properties.keys())

    @property
    def values(self) -> list[device.property]:
        """ returns the values of the properties """
        return list(self._properties.keys())

    @property
    def bluetooth(self) -> bool:
        """ Returns true if the device is used for bluetooth communication """
        return self._caption.startswith("Bluetooth")

    @property
    def guid(self) -> device.property:
        return self._class_guid

    @property
    def make(self) -> device.property:
        return self._manufacturer

    @property
    def items(self) -> list[tuple[PropertyName, property]]:
        """ returns a list[tuple[<property name>, <property value>]] or list of tuples containing the property name and the value """
        return [(name, value) for name, value in self._properties.items()]

    @property
    def properties(self) -> list[dict[device.name, device.property]] or None:
        return self._properties

    @property
    def status_info(self) -> device.property:
        return self._status_info

    @property
    def system_creation_class_name(self) -> device.property:
        return self._system_creation_class_name

    @property
    def system_name(self) -> device.property:
        return self._system_name

    @property
    def availability(self) -> device.property:
        return self._availability

    @property
    def caption(self) -> device.property:
        return self._caption

    @property
    def class_guid(self) -> device.property:
        return self._class_guid

    @property
    def compatible_id(self) -> device.property:
        return self._compatible_id

    @property
    def config_manager_error_code(self) -> device.property:
        return self._config_manager_error_code

    @property
    def config_manager_user_config(self) -> device.property:
        return self._config_manager_user_config

    @property
    def creation_class_name(self) -> device.property:
        return self._creation_class_name

    @property
    def description(self) -> device.property:
        return self._description

    @property
    def device_id(self) -> device.property:
        return self._device_id

    @property
    def error_cleared(self) -> device.property:
        return self._error_cleared

    @property
    def error_description(self) -> device.property:
        return self._error_description

    @property
    def hardware_id(self) -> device.property:
        return self._hardware_id

    @property
    def install_date(self) -> device.property:
        return self._install_date

    @property
    def last_error_code(self) -> device.property:
        return self._last_error_code

    @property
    def manufacturer(self) -> device.property:
        return self._manufacturer

    @property
    def name(self) -> device.property:
        return self._name

    @property
    def pnp_class(self) -> device.property:
        return self._pnp_class

    @property
    def pnp_device_id(self) -> device.property:
        return self._pnp_device_id

    @property
    def power_management_capabilities(self) -> device.property:
        return self._power_management_capabilities

    @property
    def power_management_support(self) -> device.property:
        return self._power_management_support

    @property
    def present(self) -> device.property:
        return self._present

    @property
    def service(self) -> device.property:
        return self._service

    @property
    def status(self) -> device.property:
        return self._status


class Devices:
    def __init__(self) -> None:
        self._devices: list[Device] = get_devices()

    def device(self, device_name: str) -> Device | None:
        """ Get a Device by Name """
        for _device in self._devices:
            if _device.name == device_name:
                return _device
            elif str(_device.name).lower() == device_name.lower():
                return _device
        return None

    def __str__(self) -> str:
        return ", ".join(self.names)

    def __int__(self) -> int:
        return len(self._devices)

    @property
    def names(self) -> list[Device.name]:
        return [_device.name for _device in self._devices]

    @property
    def properties(self) -> list[Device.properties]:
        return [_device.properties for _device in self._devices]

    @property
    def values(self) -> list[device.properties]:
        return self.properties

    @property
    def array(self) -> list[list[device.name, device.properties]]:
        return [[_device.name, [_device.properties]] for _device in self._devices]

    @property
    def device_ids(self) -> list[tuple[device.name, device.property]]:
        return [(_device.name, _device.device_id) for _device in self._devices]
