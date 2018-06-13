import platform
import re
import win32api
from collections import namedtuple
from ctypes import byref, create_unicode_buffer, windll
from ctypes.wintypes import DWORD
from itertools import count

import win32com.client
from path import Path
from tenacity import retry, wait_exponential
from winnt import PROCESSOR_ARCHITECTURE_AMD64

VER_NT_WORKSTATION = 0x0000001  # The operating system is Windows 8, Windows 7, Windows Vista, Windows XP Professional, Windows XP Home Edition, or Windows 2000 Professional.
SM_SERVERR2 = 89  # The build number if the system is Windows Server 2003 R2; otherwise, 0.

# defined at http://msdn.microsoft.com/en-us/library/aa370101(v=VS.85).aspx
UID_BUFFER_SIZE = 39
PROPERTY_BUFFER_SIZE = 256
ERROR_MORE_DATA = 234
ERROR_INVALID_PARAMETER = 87
ERROR_SUCCESS = 0
ERROR_NO_MORE_ITEMS = 259
ERROR_UNKNOWN_PRODUCT = 1605

# diff propoerties of a product, not all products have all properties
PRODUCT_PROPERTIES = ['Language',
                      'ProductName',
                      'PackageCode',
                      'Transforms',
                      'AssignmentType',
                      'PackageName',
                      'InstalledProductName',
                      'VersionString',
                      'RegCompany',
                      'RegOwner',
                      'ProductID',
                      'ProductIcon',
                      'InstallLocation',
                      'InstallSource',
                      'InstallDate',
                      'Publisher',
                      'LocalPackage',
                      'HelpLink',
                      'HelpTelephone',
                      'URLInfoAbout',
                      'URLUpdateInfo', ]

# class to be used for python users
Product = namedtuple('Product', PRODUCT_PROPERTIES)
WinName = namedtuple('WinName', ['full', 'short', 'server'])


class WinVer(object):
    def __init__(self):
        _version = platform.version().split('.')
        if len(_version) == 2:
            self.major = int(_version[0])
            self.minor = int(_version[1])
            self.build = 0
        else:
            self.major = int(_version[0])
            self.minor = int(_version[1])
            self.build = int(_version[2])
        self.wProductType = win32api.GetVersionEx(1)[8]
        self.wProcessorArchitecture = win32api.GetSystemInfo()[0]
        self.bits = 64 if platform.machine().endswith('64') else 32
        _v = self.version
        self.version_full = _v.full
        self.version_short = _v.short
        self.server = _v.server

    @property
    def dict(self):
        return {k: v for k, v in self.__dict__.items() if not k.startswith('_')}

    @property
    def version(self):
        if self.major == 5 and self.minor == 0:
            return WinName('2000', '2000', False)

        if self.major == 5 and self.minor == 1:
            return WinName('XP', 'XP', False)

        if self.major == 5 and self.minor == 2 \
                and self.wProductType == VER_NT_WORKSTATION \
                and self.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64:
            return WinName('XP Professional x64 Edition', 'XP', False)

        if self.major == 5 and self.minor == 2 \
                and win32api.GetSystemMetrics(SM_SERVERR2) == 0:
            return WinName('Server 2003', '2003', True)

        if self.major == 5 and self.minor == 2 \
                and win32api.GetSystemMetrics(SM_SERVERR2) != 0:
            return WinName('Server 2003 R2', '2003R2', True)

        if self.major == 6 and self.minor == 0 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('Vista', 'Vista', False)

        if self.major == 6 and self.minor == 0 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2008', '2008', True)

        if self.major == 6 and self.minor == 1 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2008 R2', '2008R2', True)

        if self.major == 6 and self.minor == 1 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('7', '7', False)

        if self.major == 6 and self.minor == 2 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2012', '2012', True)

        if self.major == 6 and self.minor == 2 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('8', '8', False)

        if self.major == 6 and self.minor == 3 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2012 R2', '2012R2', True)

        if self.major == 6 and self.minor == 3 \
                and self.wProductType == VER_NT_WORKSTATION:
            return WinName('8.1', '8.1', False)

        if self.major == 10 and self.minor == 0 \
                and self.wProductType != VER_NT_WORKSTATION:
            return WinName('Server 2016', '2016', True)

        if self.major == 10 and self.minor == 0 \
                and self.wProductType == VER_NT_WORKSTATION:

            if self.build <= 10240:
                return WinName('10 Threshold 1', '10', False)
            elif self.build <= 10586:
                return WinName('10 Threshold 2', '10', False)
            elif self.build <= 14393:
                return WinName('10 Redstone 1', '10RS1', False)
            elif self.build <= 15063:
                return WinName('10 Redstone 2', '10RS2', False)
            elif self.build <= 16299:
                return WinName('10 Redstone 3', '10RS3', False)
            elif self.build <= 17134:
                return WinName('10 Redstone 4', '10RS4', False)
            else:
                return WinName('10 Redstone 5', '10RS5', False)


def enum_winupdates():
    wua = win32com.client.Dispatch('Microsoft.Update.Session')
    update_searcher = wua.CreateUpdateSearcher()
    update_searcher.IncludePotentiallySupersededUpdates = True
    history_count = update_searcher.GetTotalHistoryCount()
    search_result = update_searcher.QueryHistory(0, history_count)
    updates = []
    for update in search_result:
        if update.ResultCode == 2 and update.Operation == 1:
            try:
                result = re.search(r'KB\d+', update.Title)
                updates.append(result.group(0))
            except AttributeError:
                pass
    return updates


def get_property_for_product(product, property, buf_size=PROPERTY_BUFFER_SIZE):
    """Retruns the value of a fiven property from a product."""
    property_buffer = create_unicode_buffer(buf_size)
    size = DWORD(buf_size)
    result = windll.msi.MsiGetProductInfoW(product, property, property_buffer,
                                           byref(size))
    if result == ERROR_MORE_DATA:
        return get_property_for_product(product, property,
                                        2 * buf_size)
    elif result == ERROR_SUCCESS:
        return property_buffer.value
    else:
        return None


def populate_product(uid):
    """Return a Product with the different present data."""
    properties = []
    for property in PRODUCT_PROPERTIES:
        properties.append(get_property_for_product(uid, property))
    return Product(*properties)


def get_installed_products_uids():
    """Returns a list with all the different uid of the installed apps."""
    # enum will return an error code according to the result of the app
    products = []
    for i in count(0):
        uid_buffer = create_unicode_buffer(UID_BUFFER_SIZE)
        result = windll.msi.MsiEnumProductsW(i, uid_buffer)
        if result == ERROR_NO_MORE_ITEMS:
            # done interating over the collection
            break
        products.append(uid_buffer.value)
    return products


def get_installed_products():
    """Returns a collection of products that are installed in the system."""
    products = []
    for puid in get_installed_products_uids():
        products.append(populate_product(puid))
    return products


def is_product_installed_uid(uid):
    """Return if a product with the given id is installed.

    uid Most be a unicode object with the uid of the product using
    the following format {uid}
    """
    # we try to get the VersisonString for the uid, if we get an error it means
    # that the product is not installed in the system.
    buf_size = 256
    uid_buffer = create_unicode_buffer(uid)
    property = 'VersionString'
    property_buffer = create_unicode_buffer(buf_size)
    size = DWORD(buf_size)
    result = windll.msi.MsiGetProductInfoW(uid_buffer, property, property_buffer,
                                           byref(size))
    if result == ERROR_UNKNOWN_PRODUCT:
        return False
    else:
        return True


@retry(wait=wait_exponential(multiplier=1, max=10))
def strip_hash(dir_to_strip: Path):
    # Renaming folder to the correct form
    for dir in dir_to_strip.listdir():
        old_name = dir.basename()
        new_name = re.sub(r'.[a-fA-F\d]{32}$', '', old_name)
        if old_name != new_name:
            new_path = dir_to_strip.joinpath(new_name)
            print(f'Renaming {old_name} -> {new_name}')
            if not new_path.exists():
                dir.rename(new_path)
            else:
                print(f'Directory {new_name} already exists, skipping.')
