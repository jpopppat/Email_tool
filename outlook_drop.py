import re
import struct
import tempfile
from pathlib import Path
import pythoncom
import win32clipboard
import win32com.server.util

# Windows drag-and-drop constants
DROPEFFECT_NONE = 0
DROPEFFECT_COPY = 1


class OutlookDropTarget:
    """
    Windows OLE IDropTarget for Outlook drag and drop.

    This first version accepts the drop and displays
    the available clipboard formats.
    """

    _public_methods_ = [
        "DragEnter",
        "DragOver",
        "DragLeave",
        "Drop",
    ]

    _com_interfaces_ = [
        pythoncom.IID_IDropTarget
    ]

    def __init__(self, drop_callback=None):
        self.drop_callback = drop_callback
        self.last_data_object = None

    def DragEnter(
        self,
        data_object,
        key_state,
        mouse_position,
        allowed_effect
    ):
        """
        Called when an object enters the registered window.
        """

        print("\n--- Outlook DragEnter ---")
        print(f"Allowed effect: {allowed_effect}")
        print(f"Key state: {key_state}")
        print(f"Mouse position: {mouse_position}")

        self.last_data_object = data_object

        self._print_available_formats(data_object)

        return DROPEFFECT_COPY

    def DragOver(
        self,
        key_state,
        mouse_position,
        allowed_effect
    ):
        """
        Called while the object is moved over the window.
        """

        return DROPEFFECT_COPY

    def DragLeave(self):
        """
        Called when the dragged object leaves the window.
        """

        print("--- Outlook DragLeave ---")

        self.last_data_object = None

    def Drop(
        self,
        data_object,
        key_state,
        mouse_position,
        allowed_effect
    ):
        """
        Called when the object is dropped.
        """

        print("\n--- Outlook Drop ---")
        print(f"Allowed effect: {allowed_effect}")
        print(f"Key state: {key_state}")
        print(f"Mouse position: {mouse_position}")

        self._print_available_formats(data_object)

        if self.drop_callback is not None:

            try:

                extracted_files = extract_outlook_messages(
                    data_object
                )

                print("\nExtracted Outlook files:")

                for extracted_file in extracted_files:
                    print(f"- {extracted_file}")

                self.drop_callback(
                    extracted_files
                )

            except Exception as error:

                print(
                    "Error extracting Outlook message:"
                )

                print(
                    f"{type(error).__name__}: {error}"
                )
        self.last_data_object = None

        return DROPEFFECT_COPY

    @staticmethod
    def _print_available_formats(data_object):
        """
        Display the clipboard formats offered by Outlook.
        """

        try:
            format_enumerator = data_object.EnumFormatEtc(
                pythoncom.DATADIR_GET
            )

            formats = format_enumerator.Next(100)

        except Exception as error:
            print("Could not enumerate clipboard formats:")
            print(error)
            return

        print("\nAvailable clipboard formats:")

        for format_info in formats:

            format_id = format_info[0]
            format_name = get_clipboard_format_name(
                format_id
            )

            tymed = format_info[4]
            lindex = format_info[3]

            print(
                f"- ID {format_id}: {format_name} "
                f"| tymed={tymed} "
                f"| lindex={lindex}"
)


def get_clipboard_format_name(format_id):
    """
    Return a readable Windows clipboard format name.
    """

    known_formats = {
        1: "CF_TEXT",
        2: "CF_BITMAP",
        3: "CF_METAFILEPICT",
        8: "CF_DIB",
        13: "CF_UNICODETEXT",
        15: "CF_HDROP",
        17: "CF_DIBV5",
    }

    if format_id in known_formats:
        return known_formats[format_id]

    try:
        import win32clipboard

        name = win32clipboard.GetClipboardFormatName(
            format_id
        )

        if name:
            return name

    except Exception:
        pass

    return "Unknown"

def get_available_format_names(data_object):
    """
    Return all clipboard formats provided by the drop.

    Returns:
        dict[str, int]
    """

    available_formats = {}

    format_enumerator = data_object.EnumFormatEtc(
        pythoncom.DATADIR_GET
    )

    while True:

        formats = format_enumerator.Next(50)

        if not formats:
            break

        for format_info in formats:

            format_id = format_info[0]

            format_name = get_clipboard_format_name(
                format_id
            )

            available_formats[format_name] = format_id

    return available_formats

def get_dropped_file_paths(data_object):
    """
    Read real file paths from CF_HDROP.

    Outlook or Chromium-based applications may provide
    dragged messages as real temporary files.
    """

    format_etc = (
        15,  # CF_HDROP
        None,
        pythoncom.DVASPECT_CONTENT,
        -1,
        pythoncom.TYMED_HGLOBAL
    )

    result = data_object.GetData(
        format_etc
    )

    print(
        f"CF_HDROP result type: {type(result).__name__}"
    )

    paths = []

    if isinstance(result, str):

        paths.append(
            Path(result)
        )

    elif isinstance(result, (tuple, list)):

        for item in result:

            if isinstance(item, str):

                paths.append(
                    Path(item)
                )

            elif isinstance(item, (tuple, list)):

                for nested_item in item:

                    if isinstance(nested_item, str):
                        paths.append(
                            Path(nested_item)
                        )

    if not paths:
        raise TypeError(
            "CF_HDROP wurde empfangen, aber es konnten "
            "keine Dateipfade gelesen werden. "
            f"Erhaltener Datentyp: {type(result).__name__}"
        )

    return paths

def sanitize_windows_filename(filename):
    """
    Remove characters that are invalid in Windows filenames.
    """

    filename = str(filename).strip()

    filename = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        filename
    )

    filename = filename.rstrip(". ")

    if not filename:
        filename = "Outlook_Email.msg"

    return filename

def get_outlook_filenames(data_object):
    """
    Read filenames from Outlook's FileGroupDescriptorW format.

    Returns:
        list[str]
    """

    format_id = win32clipboard.RegisterClipboardFormat(
        "FileGroupDescriptorW"
    )

    format_etc = (
        format_id,
        None,
        pythoncom.DVASPECT_CONTENT,
        -1,
        pythoncom.TYMED_HGLOBAL
    )

    descriptor_medium = data_object.GetData(
        format_etc
    )

    print(
        f"FileGroupDescriptorW medium: "
        f"{type(descriptor_medium).__name__}"
    )

    # pywin32 returns a PySTGMEDIUM.
    # Its .data property contains the actual HGLOBAL data.
    descriptor_data = descriptor_medium.data

    print(
        f"Descriptor data type: "
        f"{type(descriptor_data).__name__}"
    )

    if isinstance(descriptor_data, str):
        descriptor_data = descriptor_data.encode(
            "latin1"
        )

    if not isinstance(
        descriptor_data,
        (bytes, bytearray)
    ):
        raise TypeError(
            "Unexpected FileGroupDescriptorW data type: "
            f"{type(descriptor_data).__name__}"
        )

    descriptor_data = bytes(
        descriptor_data
    )

    descriptor_data = bytes(descriptor_data)

    # First DWORD contains the number of virtual files
    file_count = struct.unpack_from(
        "<I",
        descriptor_data,
        0
    )[0]

    filenames = []

    # FILEDESCRIPTORW is 592 bytes.
    # The filename starts at byte 72 within each descriptor.
    descriptor_size = 592
    filename_offset = 72
    filename_length = 520  # 260 UTF-16 characters

    for index in range(file_count):

        descriptor_start = (
            4 + index * descriptor_size
        )

        filename_start = (
            descriptor_start + filename_offset
        )

        filename_end = (
            filename_start + filename_length
        )

        raw_filename = descriptor_data[
            filename_start:filename_end
        ]

        filename = raw_filename.decode(
            "utf-16-le",
            errors="ignore"
        )

        filename = filename.split(
            "\x00",
            1
        )[0]

        filename = sanitize_windows_filename(
            filename
        )

        if not filename.lower().endswith(".msg"):
            filename += ".msg"

        filenames.append(filename)

    return filenames

def save_outlook_file_content(
    data_object,
    file_index,
    destination
):
    """
    Save one virtual Outlook file from FileContents.

    Outlook may provide FileContents as:
    - IStream
    - IStorage
    - or both

    Different Outlook versions may require lindex=-1 or
    the actual file index.
    """

    format_id = win32clipboard.RegisterClipboardFormat(
        "FileContents"
    )

    destination = Path(destination)

    requests = [
        # Use exactly what Outlook advertised first
        (
            -1,
            pythoncom.TYMED_ISTREAM
            | pythoncom.TYMED_ISTORAGE
        ),

        # Some Outlook versions require the file index
        (
            file_index,
            pythoncom.TYMED_ISTREAM
            | pythoncom.TYMED_ISTORAGE
        ),

        # Individual fallback formats
        (-1, pythoncom.TYMED_ISTREAM),
        (file_index, pythoncom.TYMED_ISTREAM),

        (-1, pythoncom.TYMED_ISTORAGE),
        (file_index, pythoncom.TYMED_ISTORAGE),
    ]

    last_error = None
    result = None

    for lindex, tymed in requests:

        format_etc = (
            format_id,
            None,
            pythoncom.DVASPECT_CONTENT,
            lindex,
            tymed
        )

        print(
            f"Requesting FileContents: "
            f"index={lindex}, tymed={tymed}"
        )

        try:
            result = data_object.GetData(
                format_etc
            )

            print(
                "FileContents request successful."
            )

            break

        except pythoncom.com_error as error:

            last_error = error

            print(
                f"FileContents request failed: "
                f"{error}"
            )

    if result is None:

        raise RuntimeError(
            "FileContents konnte mit keinem unterstützten "
            "FORMATETC abgerufen werden."
        ) from last_error

    return save_stgmedium_to_file(
        result,
        destination
    )
def save_stgmedium_to_file(
    medium,
    destination
):
    """
    Save the contents of a PySTGMEDIUM to a file.

    Supports:
    - bytes / HGLOBAL
    - IStream
    - IStorage
    """

    destination = Path(destination)

    print(
        f"Received medium type: "
        f"{type(medium).__name__}"
    )

    # Get actual data from PySTGMEDIUM
    if hasattr(medium, "data"):

        print(
            f"STGMEDIUM tymed: "
            f"{getattr(medium, 'tymed', 'unknown')}"
        )

        data = medium.data

    else:
        data = medium

    print(
        f"Medium data type: "
        f"{type(data).__name__}"
    )


    # =====================================================
    # BYTE DATA
    # =====================================================

    if isinstance(data, (bytes, bytearray)):

        destination.write_bytes(
            bytes(data)
        )

        return destination


    # =====================================================
    # ISTREAM
    # =====================================================

    if hasattr(data, "Read"):

        with open(destination, "wb") as output_file:

            while True:

                chunk = data.Read(
                    1024 * 1024
                )

                if not chunk:
                    break

                output_file.write(chunk)

        return destination


    # =====================================================
    # ISTORAGE
    # =====================================================

    if hasattr(data, "CopyTo"):

        return save_istorage_to_msg(
            data,
            destination
        )


    raise TypeError(
        "Nicht unterstützter FileContents-Datentyp: "
        f"{type(data).__name__}"
    )

def save_istorage_to_msg(
    source_storage,
    destination
):
    """
    Copy an Outlook IStorage object into a new MSG file.
    """

    destination = Path(destination)

    if destination.exists():
        destination.unlink()

    # Windows STGM constants
    STGM_READWRITE = 0x00000002
    STGM_SHARE_EXCLUSIVE = 0x00000010
    STGM_CREATE = 0x00001000

    storage_flags = (
        STGM_CREATE
        | STGM_READWRITE
        | STGM_SHARE_EXCLUSIVE
    )

    destination_storage = pythoncom.StgCreateDocfile(
        str(destination),
        storage_flags,
        0
    )

    try:
        # PyIStorage.CopyTo expects exactly:
        # exclude_iids, exclude_names, destination_storage
        source_storage.CopyTo(
            None,
            None,
            destination_storage
        )

        # STGC_DEFAULT = 0
        destination_storage.Commit(0)

    finally:
        # Release the destination COM storage
        destination_storage = None

    if not destination.exists():
        raise FileNotFoundError(
            f"MSG-Datei wurde nicht erstellt: {destination}"
        )

    if destination.stat().st_size == 0:
        destination.unlink(missing_ok=True)

        raise ValueError(
            "Die erzeugte MSG-Datei ist leer."
        )

    print(
        f"IStorage successfully saved: "
        f"{destination}"
    )

    return destination

def extract_outlook_messages(data_object):
    """
    Extract dropped Outlook messages.

    Supports:
    1. CF_HDROP real files
    2. Classic Outlook virtual files using
       FileGroupDescriptorW and FileContents

    Returns:
        list[Path]
    """

    formats = get_available_format_names(
        data_object
    )

    print("\nDetected drop formats:")

    for format_name in formats:
        print(f"- {format_name}")


    # =====================================================
    # CASE 1: REAL FILES PROVIDED THROUGH CF_HDROP
    # =====================================================

    if "CF_HDROP" in formats:

        print("\nUsing CF_HDROP file drop.")

        dropped_paths = get_dropped_file_paths(
            data_object
        )

        valid_files = []

        for path in dropped_paths:

            path = Path(path)

            print(f"Dropped file: {path}")

            if path.exists() and path.is_file():
                valid_files.append(path)

        if not valid_files:
            raise FileNotFoundError(
                "CF_HDROP enthielt keine erreichbaren Dateien."
            )

        return valid_files


    # =====================================================
    # CASE 2: CLASSIC OUTLOOK VIRTUAL FILE
    # =====================================================

    has_descriptor = (
        "FileGroupDescriptorW" in formats
        or "FileGroupDescriptor" in formats
    )

    if has_descriptor and "FileContents" in formats:

        print(
            "\nUsing Outlook virtual file extraction."
        )

        filenames = get_outlook_filenames(
            data_object
        )

        if not filenames:
            raise ValueError(
                "Outlook did not provide any filenames."
            )

        temporary_root = (
            Path(tempfile.gettempdir())
            / "EmailTool"
        )

        temporary_root.mkdir(
            parents=True,
            exist_ok=True
        )

        extracted_files = []

        for index, filename in enumerate(filenames):

            destination = (
                temporary_root
                / filename
            )

            if destination.exists():

                counter = 2

                while True:

                    candidate = (
                        destination.parent
                        / (
                            f"{destination.stem}_{counter}"
                            f"{destination.suffix}"
                        )
                    )

                    if not candidate.exists():
                        destination = candidate
                        break

                    counter += 1

            saved_file = save_outlook_file_content(
                data_object=data_object,
                file_index=index,
                destination=destination
            )

            extracted_files.append(
                saved_file
            )

        return extracted_files


    # =====================================================
    # UNSUPPORTED DROP
    # =====================================================

    format_list = ", ".join(
        sorted(formats.keys())
    )

    raise ValueError(
        "Der Drop enthält weder verwendbare CF_HDROP-Dateien "
        "noch FileGroupDescriptor/FileContents.\n\n"
        f"Erkannte Formate: {format_list}"
    )


class OutlookDropManager:
    """
    Register and unregister an Outlook-compatible
    Windows OLE drop target for a Tkinter window.
    """

    def __init__(self, tkinter_window, drop_callback=None):
        self.window = tkinter_window
        self.drop_callback = drop_callback

        self.hwnd = None
        self.drop_target = None
        self.wrapped_drop_target = None
        self.registered = False

    def register(self):
        """
        Register the Tkinter window as a Windows OLE drop target.
        """

        if self.registered:
            return

        # Ensure the native Tk window has been created
        self.window.update_idletasks()
        self.window.update()

        self.hwnd = self.window.winfo_id()

        self.drop_target = OutlookDropTarget(
            drop_callback=self.drop_callback
        )

        # Wrap the Python object as a COM object
        self.wrapped_drop_target = (
            win32com.server.util.wrap(
                self.drop_target,
                pythoncom.IID_IDropTarget
            )
        )

        try:
            pythoncom.RegisterDragDrop(
                self.hwnd,
                self.wrapped_drop_target
            )

        except pythoncom.com_error as error:

            # DRAGDROP_E_ALREADYREGISTERED
            if error.hresult == -2147221247:
                raise RuntimeError(
                    "Das Fenster ist bereits als "
                    "Drag-and-drop-Ziel registriert."
                ) from error

            raise

        self.registered = True

        print("\nOutlook OLE drop target registered.")
        print(f"Window handle: {self.hwnd}")

    def unregister(self):
        """
        Remove the Windows OLE drop target registration.
        """

        if not self.registered:
            return

        try:
            pythoncom.RevokeDragDrop(
                self.hwnd
            )

        except pythoncom.com_error as error:
            print(
                "Could not revoke Outlook drop target:"
            )
            print(error)

        self.registered = False
        self.drop_target = None
        self.wrapped_drop_target = None

        print("Outlook OLE drop target unregistered.")