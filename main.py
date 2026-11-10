from pathlib import Path
import extract_msg
import yaml
import re
from datetime import datetime
import shutil
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
from tkinterdnd2 import DND_FILES, TkinterDnD

def read_email(msg_path):
    """
    Read an Outlook MSG file and separate real attachments
    from inline/signature images.
    """

    msg_path = Path(msg_path)

    if not msg_path.exists():
        raise FileNotFoundError(f"MSG file not found: {msg_path}")

    if msg_path.suffix.lower() != ".msg":
        raise ValueError(f"File is not an MSG file: {msg_path}")

    msg = extract_msg.Message(str(msg_path))

    # ----- HTML body -----

    html_body = msg.htmlBody or b""

    if isinstance(html_body, bytes):
        html_body = html_body.decode("utf-8", errors="ignore")

    html_lower = html_body.lower()


    # ----- Basic email data -----

    email_data = {
        "subject": msg.subject or "",
        "sender": msg.sender or "",
        "to": msg.to or "",
        "cc": msg.cc or "",
        "date": msg.date or "",
        "body": msg.body or "",
        "html_body": html_body,
        "attachments": [],
        "inline_attachments": [],
        "source_path": msg_path,

        # Keep MSG object alive while attachments are being used
        "_msg": msg,
    }


    # ----- Attachments -----

    for attachment in msg.attachments:

        filename = (
            getattr(attachment, "longFilename", None)
            or getattr(attachment, "shortFilename", None)
            or "unnamed_attachment"
        )

        content_id = getattr(attachment, "cid", None)

        if content_id:
            content_id = str(content_id).strip("<>")


        # ----- Detect inline attachment -----

        is_inline = False
        reason = None


        # 1. Content-ID referenced by HTML
        if content_id:
            if f"cid:{content_id}".lower() in html_lower:
                is_inline = True
                reason = "CID referenced in HTML"


 


        # ----- Create attachment data -----

        attachment_data = {
            "filename": filename,
            "content_id": content_id,
            "is_inline": is_inline,
            "inline_reason": reason,
            "attachment": attachment,
        }


        # ----- Sort -----

        if is_inline:
            email_data["inline_attachments"].append(
                attachment_data
            )

        else:
            email_data["attachments"].append(
                attachment_data
            )


    return email_data

def scan_projects(root_folder, config):
    """
    Scan the project directory.

    Expected structure:
        ROOT/
            YYYY/
                YY_XXX_Project Name/

    Only folders matching the project naming pattern are considered.
    All other folders are silently ignored.
    """

    root_folder = Path(root_folder)

    if not root_folder.exists():
        raise FileNotFoundError(
            f"Root folder not found: {root_folder}"
        )

   

    projects = []

    # Example:
    # 26_001_Project Name
    project_pattern = re.compile(
        r"^(?P<prefix>A?)(?P<year>\d{2})_(?P<number>\d{3})_(?P<name>.+)$"
    )
    # Scan root folder
    try:
        year_folders = list(root_folder.iterdir())

    except PermissionError:
        raise PermissionError(
            f"No permission to access root folder: {root_folder}"
        )
  # Locations that contain YYYY folders
    search_roots = [
        root_folder / folder
        for folder in config["paths"]["project_search_folders"]
    ]

    for search_root in search_roots:

        if not search_root.exists():
            continue

        try:
            year_folders = list(search_root.iterdir())
        except PermissionError:
            continue
        for year_folder in year_folders:

            # First check the name.
            # This does not require entering the folder.
            if not re.fullmatch(r"\d{4}", year_folder.name):
                continue

            try:
                if not year_folder.is_dir():
                    continue

                project_folders = list(year_folder.iterdir())

            except PermissionError:
                # Skip inaccessible year folders
                continue

            for project_folder in project_folders:

                # IMPORTANT:
                # Check project name BEFORE trying to access the folder.
                match = project_pattern.fullmatch(project_folder.name)

                if not match:
                    continue

                try:
                    if not project_folder.is_dir():
                        continue

                except PermissionError:
                    # Valid-looking project, but user has no access
                    continue

                projects.append({
                    "year": year_folder.name,

                    "project_number": (
                        f"{match.group('prefix')}"
                        f"{match.group('year')}_"
                        f"{match.group('number')}"
                    ),

                    "name": match.group("name"),

                    "folder_name": project_folder.name,

                    "path": project_folder,
                })

    # Sort by year and project number
    projects.sort(
        key=lambda project: (
            project["year"],
            project["project_number"]
        )
    )

    return projects

def search_projects(projects, search_term):
    """
    Search projects by project number or project name.

    Examples:
        "26_001"
        "26001"
        "001"
        "Fusch"
        "wasserversorgung"

    Returns a list of matching projects.
    """

    search_term = search_term.strip().lower()

    if not search_term:
        return []

    # Version without "_" for number searches
    normalized_search = search_term.replace("_", "").replace(" ", "")

    results = []

    for project in projects:

        project_number = project["project_number"].lower()
        project_name = project["name"].lower()

        normalized_number = project_number.replace("_", "")

        # Search project number
        number_match = (
            search_term in project_number
            or normalized_search in normalized_number
        )

        # Search project name
        name_match = search_term in project_name

        if number_match or name_match:
            results.append(project)

    return results

def load_config():
    """Load settings from config.yaml."""

    config_path = Path(__file__).parent / "config.yaml"

    with open(config_path, "r", encoding="utf-8") as file:
        config = yaml.safe_load(file)

    return config

def sanitize_filename(name):
    """
    Remove characters that are not allowed in Windows file or folder names.
    """

    name = str(name).strip()

    name = re.sub(r'[<>:"/\\|?*]', "_", name)

    name = name.rstrip(". ")

    if not name:
        return "Unbenannt"

    return name

def create_attachment_destination(
    name1_folder,
    name2,
    email_date=None
):
    """
    Create the dated attachment folder:

        name1/YYYY_MM_DD_name2
    """

    name2 = sanitize_filename(name2)

    if email_date:
        try:
            folder_date = email_date.strftime(
                "%Y_%m_%d"
            )
        except AttributeError:
            folder_date = datetime.now().strftime(
                "%Y_%m_%d"
            )
    else:
        folder_date = datetime.now().strftime(
            "%Y_%m_%d"
        )

    folder_name = f"{folder_date}_{name2}"

    destination = (
        Path(name1_folder)
        / folder_name
    )

    destination.mkdir(
        parents=True,
        exist_ok=True
    )

    return destination


def get_unique_path(file_path):
    """
    Return a unique file path without overwriting existing files.

    Example:
        Plan.pdf
        Plan_2.pdf
        Plan_3.pdf
    """

    file_path = Path(file_path)

    if not file_path.exists():
        return file_path

    counter = 2

    while True:
        new_path = (
            file_path.parent
            / f"{file_path.stem}_{counter}{file_path.suffix}"
        )

        if not new_path.exists():
            return new_path

        counter += 1

def select_email_direction():
    """
    Ask whether the email is incoming or outgoing.

    Returns:
        "Ein" for incoming email
        "Aus" for outgoing email
    """

    while True:
        print("\n--- E-MAIL ABLAGE ---")
        print("1: Ein  (Eingang)")
        print("2: Aus  (Ausgang)")

        selection = input(
            "\nWo soll die E-Mail gespeichert werden? "
        ).strip().lower()

        if selection in {"1", "ein", "eingang"}:
            return "ein"

        if selection in {"2", "aus", "ausgang"}:
            return "aus"

        print("Ungültige Auswahl.")

def save_email(
    email,
    selected_project,
    config,
    email_direction
):
    """
    Copy the original MSG file into:

        project/email_folder/Ein
    or
        project/email_folder/Aus

    Filename:
        YYYY_MM_DD_originalname.msg
    """

    source_path = Path(email["source_path"])
    sender = email["sender"].split("<")[0].strip()
    sender = sanitize_filename(sender)

    
    email_folder = (
        Path(selected_project["path"])
        / config["paths"]["email_folder"]
        / email_direction
    )

    email_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    # Email date
    email_date = email["date"]

    try:
        date_string = email_date.strftime("%Y_%m_%d")
    except (AttributeError, ValueError):
        date_string = "unknown_date"

    original_name = sanitize_filename(
        source_path.stem
    )

    filename = (
        f"Em_{date_string}_{sender}_{original_name}.msg"
    )

    destination = email_folder / filename

    destination = get_unique_path(destination)

    shutil.copy2(
        source_path,
        destination
    )

    return destination

def save_attachments(
    selected_attachments,
    destination_folder
):
    """
    Save selected MSG attachments.

    Existing files will not be overwritten.
    """

    destination_folder = Path(destination_folder)

    destination_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    saved_files = []

    for attachment_data in selected_attachments:

        attachment = attachment_data["attachment"]

        filename = sanitize_filename(
            attachment_data["filename"]
        )

        destination = (
            destination_folder
            / filename
        )

        destination = get_unique_path(
            destination
        )

        attachment.save(
            customPath=str(destination_folder),
            customFilename=destination.name
        )

        saved_files.append(destination)

    return saved_files

def load_user_settings():
    """
    Load user-specific settings.
    """

    settings_path = (
        Path(__file__).parent
        / "user_settings.json"
    )

    if not settings_path.exists():
        return {
            "recent_projects": []
        }

    try:
        with open(
            settings_path,
            "r",
            encoding="utf-8"
        ) as file:

            return json.load(file)

    except (json.JSONDecodeError, OSError):
        return {
            "recent_projects": []
        }

def save_user_settings(settings):
    """
    Save user-specific settings.
    """

    settings_path = (
        Path(__file__).parent
        / "user_settings.json"
    )

    with open(
        settings_path,
        "w",
        encoding="utf-8"
    ) as file:

        json.dump(
            settings,
            file,
            indent=4,
            ensure_ascii=False
        )


''' alte Funktionen für terminalauswahl, werden jetzt durch GUI ersetzt

def select_name1_folder(project_path, config):
    """
    Let the user select an existing name1 folder or create a new one.
    """

    attachment_root = (
        Path(project_path)
        / config["paths"]["attachment_folder"]
    )

    attachment_root.mkdir(
        parents=True,
        exist_ok=True
    )

    existing_folders = sorted(
        [
            folder
            for folder in attachment_root.iterdir()
            if folder.is_dir()
        ],
        key=lambda folder: folder.name.lower()
    )

    print("\n--- ORGANISATION / FIRMA / GEMEINDE ---")

    if existing_folders:

        for index, folder in enumerate(
            existing_folders,
            start=1
        ):
            print(f"{index}: {folder.name}")

    else:
        print("Noch keine vorhandenen Ordner.")

    print("\nN: Neuen Ordner anlegen")

    selection = input(
        "\nNummer oder neuer Name: "
    ).strip()

    if selection.isdigit():

        index = int(selection) - 1

        if 0 <= index < len(existing_folders):
            return existing_folders[index]

        print("Ungültige Auswahl.")

    name1 = sanitize_filename(selection)

    name1_folder = attachment_root / name1

    name1_folder.mkdir(
        parents=True,
        exist_ok=True
    )

    return name1_folder

    
def select_attachments(attachments):
    """
    Let the user choose which attachments should be saved.

    Input examples:
        A
        1
        1,3,4
        N
    """

    if not attachments:
        print("\nKeine echten Anhänge vorhanden.")
        return []

    print("\n--- ANHÄNGE ---")

    for index, attachment in enumerate(
        attachments,
        start=1
    ):
        print(
            f"{index}: "
            f"{attachment['filename']}"
        )

    print("\nA: Alle Anhänge")
    print("N: Keine Anhänge")
    print("Oder Nummern eingeben, z. B. 1,3,4")

    selection = input(
        "\nWelche Anhänge sollen gespeichert werden? "
    ).strip().lower()

    if selection == "a":
        return attachments

    if selection in {"n", ""}:
        return []

    selected_attachments = []

    for value in selection.split(","):

        value = value.strip()

        if not value.isdigit():
            print(
                f"Ungültige Auswahl wird ignoriert: {value}"
            )
            continue

        index = int(value) - 1

        if 0 <= index < len(attachments):

            attachment = attachments[index]

            if attachment not in selected_attachments:
                selected_attachments.append(
                    attachment
                )

        else:
            print(
                f"Anhangnummer außerhalb des Bereichs: "
                f"{value}"
            )

    return selected_attachments



def select_project(projects):
    """
    Search for a project and let the user select one.

    Returns:
        dict: Selected project
        None: User cancelled the selection
    """

    while True:

        search_term = input(
            "\nSearch project (Enter to cancel): "
        ).strip()

        if not search_term:
            return None

        results = search_projects(
            projects,
            search_term
        )

        if not results:
            print("No projects found.")
            continue

        print(f"\n{len(results)} project(s) found:")

        for index, project in enumerate(results, start=1):
            print(
                f"{index}: "
                f"{project['project_number']} | "
                f"{project['name']}"
            )

        # Exactly one result
        if len(results) == 1:

            selection = input(
                "\nPress Enter to select this project "
                "or 'B' to go back: "
            ).strip().lower()

            if selection == "b":
                continue

            return results[0]

        # Multiple results
        selection = input(
            "\nSelect project number "
            "(B = back, Q = cancel): "
        ).strip().lower()

        if selection == "q":
            return None

        if selection == "b":
            continue

        if not selection.isdigit():
            print("Invalid selection.")
            continue

        index = int(selection) - 1

        if not 0 <= index < len(results):
            print("Invalid selection.")
            continue

        return results[index]

'''
# ==========================================
# GUI
# ==========================================

class EmailToolGUI:

    def __init__(self, root, config, projects):
        self.root = root
        self.config = config
        self.projects = projects

        self.current_results = projects
        self.selected_project = None

        self.user_settings = load_user_settings()
        self.recent_projects = []

        recent_paths = self.user_settings.get(
            "recent_projects",
            []
        )

        for recent_path in recent_paths:

            for project in self.projects:

                if str(project["path"]) == recent_path:
                    self.recent_projects.append(project)
                    break

        self.email = None
        self.attachment_vars = []
        self.email_direction = tk.StringVar(
            value="ein"
        )
        self.name1_var = tk.StringVar()
        self.name2_var = tk.StringVar()
        self.year_var = tk.StringVar(value="Alle Jahre")
        

        self.root.title("Email Tool")
        self.root.geometry("1000x700")
        self.root.minsize(850, 600)


        self.create_widgets()

    ''' alt eine spalte
        def create_widgets(self):

            # -------------------------
            # Title
            # -------------------------
            title = ttk.Label(
                self.root,
                text="Email Tool",
                font=("Segoe UI", 18, "bold")
            )
            title.pack(pady=(20, 10))




            # -------------------------
            # Project section
            # -------------------------

            project_frame = ttk.LabelFrame(
                self.root,
                text="Projekt auswählen",
                padding=10
            )
            project_frame.pack(
                fill="both",
                expand=True,
                padx=20,
                pady=10
            )
            # -------------------------
            # Recent projects
            # -------------------------

            ttk.Label(
                project_frame,
                text="Zuletzt verwendet:"
            ).pack(
                anchor="w",
                pady=(0, 5)
            )

            self.recent_projects_frame = ttk.Frame(
                project_frame
            )

            self.recent_projects_frame.pack(
                fill="x",
                pady=(0, 10)
            )

            self.update_recent_projects()
            # -------------------------
            # Year filter
            # -------------------------

            year_frame = ttk.Frame(
                project_frame
            )

            year_frame.pack(
                fill="x",
                pady=(0, 10)
            )

            ttk.Label(
                year_frame,
                text="Projektjahr:"
            ).pack(
                side="left"
            )

            years = sorted(
                {project["year"] for project in self.projects},
                reverse=True
            )

            self.year_combobox = ttk.Combobox(
                year_frame,
                textvariable=self.year_var,
                values=["Alle Jahre"] + years,
                state="readonly",
                width=15
            )

            self.year_combobox.pack(
                side="left",
                padx=(10, 0)
            )

            self.year_combobox.bind(
                "<<ComboboxSelected>>",
                self.update_project_list
            )

            # Search
            ttk.Label(
                project_frame,
                text="Projekt suchen:"
            ).pack(anchor="w")

            self.search_var = tk.StringVar()

            self.search_entry = ttk.Entry(
                project_frame,
                textvariable=self.search_var
            )
            self.search_entry.pack(
                fill="x",
                pady=(5, 10)
            )

            self.search_entry.bind(
                "<KeyRelease>",
                self.update_project_list
            )


            # Project list
            project_list_frame = ttk.Frame(
                project_frame
            )

            project_list_frame.pack(
                fill="both",
                expand=True
            )
            self.project_list = tk.Listbox(
                project_list_frame,
                height=7
            )

            self.project_list.pack(
                side="left",
                fill="both",
                expand=True
            )

            project_scrollbar = ttk.Scrollbar(
                project_list_frame,
                orient="vertical",
                command=self.project_list.yview
            )

            project_scrollbar.pack(
                side="right",
                fill="y"
            )

            self.project_list.config(
                yscrollcommand=project_scrollbar.set
            )

            self.project_list.bind(
                "<<ListboxSelect>>",
                self.on_project_selected
            )


            # Selected project
            self.selected_project_label = ttk.Label(
                project_frame,
                text="Kein Projekt ausgewählt."
            )
            self.selected_project_label.pack(
                anchor="w",
                pady=(10, 0)
            )



            # -------------------------
            # Email section
            # -------------------------

            email_frame = ttk.LabelFrame(
                self.root,
                text="E-Mail",
                padding=10
            )

            email_frame.pack(
                fill="x",
                padx=20,
                pady=10
            )


            # Select MSG button
            self.select_email_button = ttk.Button(
                email_frame,
                text="MSG auswählen",
                command=self.select_email
            )

            self.select_email_button.pack(
                anchor="w",
                pady=(0, 10)
            )


            # Email information
            self.email_subject_label = ttk.Label(
                email_frame,
                text="Betreff: -"
            )

            self.email_subject_label.pack(
                anchor="w"
            )


            self.email_sender_label = ttk.Label(
                email_frame,
                text="Absender: -"
            )

            self.email_sender_label.pack(
                anchor="w"
            )


            self.email_date_label = ttk.Label(
                email_frame,
                text="Datum: -"
            )

            self.email_date_label.pack(
                anchor="w"
            )


            self.email_attachment_label = ttk.Label(
                email_frame,
                text="Anhänge: -"
            )

            self.email_attachment_label.pack(
                anchor="w"
            )

            # -------------------------
            # Email direction
            # -------------------------

            direction_frame = ttk.Frame(
                email_frame
            )

            direction_frame.pack(
                fill="x",
                pady=(10, 0)
            )

            ttk.Label(
                direction_frame,
                text="Ablage:"
            ).pack(
                side="left"
            )

            ttk.Radiobutton(
                direction_frame,
                text="Eingang",
                variable=self.email_direction,
                value="ein"
            ).pack(
                side="left",
                padx=(10, 5)
            )

            ttk.Radiobutton(
                direction_frame,
                text="Ausgang",
                variable=self.email_direction,
                value="aus"
            ).pack(
                side="left",
                padx=5
            )

            # -------------------------
            # Attachment section
            # -------------------------

            self.attachment_frame = ttk.LabelFrame(
                self.root,
                text="Anhänge",
                padding=10
            )

            self.attachment_frame.pack(
                fill="x",
                padx=20,
                pady=10
            )

            self.no_attachments_label = ttk.Label(
                self.attachment_frame,
                text="Noch keine E-Mail ausgewählt."
            )

            self.no_attachments_label.pack(
                anchor="w"
            )

            # -------------------------
            # Attachment destination
            # -------------------------

            self.attachment_destination_frame = ttk.LabelFrame(
                self.root,
                text="Ablage der Anhänge",
                padding=10
            )

            self.attachment_destination_frame.pack(
                fill="x",
                padx=20,
                pady=10
            )


            # name1
            ttk.Label(
                self.attachment_destination_frame,
                text="Gemeinde / Organisation / Firma:"
            ).pack(
                anchor="w"
            )


            self.name1_combobox = ttk.Combobox(
                self.attachment_destination_frame,
                textvariable=self.name1_var
            )

            self.name1_combobox.pack(
                fill="x",
                pady=(5, 0)
            )

            # name2
            ttk.Label(
                self.attachment_destination_frame,
                text="Art der Unterlagen:"
            ).pack(
                anchor="w",
                pady=(10, 0)
            )

            self.name2_entry = ttk.Entry(
                self.attachment_destination_frame,
                textvariable=self.name2_var
            )

            self.name2_entry.pack(
                fill="x",
                pady=(5, 0)
            )

            # -------------------------
            # Save button
            # -------------------------

            self.save_button = ttk.Button(
                self.root,
                text="Ablegen",
                command=self.save_email_and_attachments
            )

            self.save_button.pack(
                pady=20
            )



            # Display projects initially
            self.update_project_list()
    '''
            
    def create_widgets(self):

        # =====================================================
        # MAIN CONTAINER
        # =====================================================

        main_frame = ttk.Frame(
            self.root,
            padding=15
        )

        main_frame.pack(
            fill="both",
            expand=True
        )

        # Two responsive columns
        main_frame.columnconfigure(
            0,
            weight=1
        )

        main_frame.columnconfigure(
            1,
            weight=1
        )

        main_frame.rowconfigure(
            1,
            weight=1
        )


        # =====================================================
        # TITLE
        # =====================================================

        title = ttk.Label(
            main_frame,
            text="Email Tool",
            font=("Segoe UI", 18, "bold")
        )

        title.grid(
            row=0,
            column=0,
            columnspan=2,
            pady=(0, 15)
        )


        # =====================================================
        # LEFT COLUMN
        # =====================================================

        left_frame = ttk.Frame(
            main_frame
        )

        left_frame.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=(0, 8)
        )


        # =====================================================
        # EMAIL SECTION
        # =====================================================

        email_frame = ttk.LabelFrame(
            left_frame,
            text="E-Mail",
            padding=10
        )

        email_frame.pack(
            fill="x",
            pady=(0, 10)
        )


        # Select MSG button

        self.select_email_button = ttk.Button(
            email_frame,
            text="MSG auswählen",
            command=self.select_email
        )

        self.select_email_button.pack(
            anchor="w",
            pady=(0, 10)
        )

        self.drop_label = ttk.Label(
            email_frame,
            text="MSG-Datei hier hineinziehen",
            anchor="center",
            padding=15
        )

        self.drop_label.pack(
            fill="x",
            pady=(0, 10)
        )

        self.drop_label.drop_target_register(
            DND_FILES
        )

        self.drop_label.dnd_bind(
            "<<Drop>>",
            self.on_email_drop
        )

        # Email information

        self.email_subject_label = ttk.Label(
            email_frame,
            text="Betreff: -"
        )

        self.email_subject_label.pack(
            anchor="w"
        )


        self.email_sender_label = ttk.Label(
            email_frame,
            text="Absender: -"
        )

        self.email_sender_label.pack(
            anchor="w"
        )


        self.email_date_label = ttk.Label(
            email_frame,
            text="Datum: -"
        )

        self.email_date_label.pack(
            anchor="w"
        )


        self.email_attachment_label = ttk.Label(
            email_frame,
            text="Anhänge: -"
        )

        self.email_attachment_label.pack(
            anchor="w"
        )


        # =====================================================
        # EMAIL DIRECTION
        # =====================================================

        direction_frame = ttk.Frame(
            email_frame
        )

        direction_frame.pack(
            fill="x",
            pady=(10, 0)
        )


        ttk.Label(
            direction_frame,
            text="Ablage:"
        ).pack(
            side="left"
        )


        ttk.Radiobutton(
            direction_frame,
            text="Eingang",
            variable=self.email_direction,
            value="ein"
        ).pack(
            side="left",
            padx=(10, 5)
        )


        ttk.Radiobutton(
            direction_frame,
            text="Ausgang",
            variable=self.email_direction,
            value="aus"
        ).pack(
            side="left",
            padx=5
        )


        # =====================================================
        # ATTACHMENTS
        # =====================================================

        self.attachment_frame = ttk.LabelFrame(
            left_frame,
            text="Anhänge",
            padding=10
        )

        self.attachment_frame.pack(
            fill="x",
            pady=(0, 10)
        )


        self.no_attachments_label = ttk.Label(
            self.attachment_frame,
            text="Noch keine E-Mail ausgewählt."
        )

        self.no_attachments_label.pack(
            anchor="w"
        )


        # =====================================================
        # ATTACHMENT DESTINATION
        # =====================================================

        self.attachment_destination_frame = ttk.LabelFrame(
            left_frame,
            text="Ablage der Anhänge",
            padding=10
        )

        self.attachment_destination_frame.pack(
            fill="x",
            pady=(0, 10)
        )


        # Name 1

        ttk.Label(
            self.attachment_destination_frame,
            text="Gemeinde / Organisation / Firma:"
        ).pack(
            anchor="w"
        )


        self.name1_combobox = ttk.Combobox(
            self.attachment_destination_frame,
            textvariable=self.name1_var
        )

        self.name1_combobox.pack(
            fill="x",
            pady=(5, 0)
        )


        # Name 2

        ttk.Label(
            self.attachment_destination_frame,
            text="Art der Unterlagen:"
        ).pack(
            anchor="w",
            pady=(10, 0)
        )


        self.name2_entry = ttk.Entry(
            self.attachment_destination_frame,
            textvariable=self.name2_var
        )

        self.name2_entry.pack(
            fill="x",
            pady=(5, 0)
        )


        # =====================================================
        # RIGHT COLUMN
        # =====================================================

        right_frame = ttk.Frame(
            main_frame
        )

        right_frame.grid(
            row=1,
            column=1,
            sticky="nsew",
            padx=(8, 0)
        )


        # =====================================================
        # PROJECT SECTION
        # =====================================================

        project_frame = ttk.LabelFrame(
            right_frame,
            text="Projekt auswählen",
            padding=10
        )

        project_frame.pack(
            fill="both",
            expand=True
        )


        # =====================================================
        # RECENT PROJECTS
        # =====================================================

        ttk.Label(
            project_frame,
            text="Zuletzt verwendet:"
        ).pack(
            anchor="w",
            pady=(0, 5)
        )


        self.recent_projects_frame = ttk.Frame(
            project_frame
        )

        self.recent_projects_frame.pack(
            fill="x",
            pady=(0, 10)
        )


        self.update_recent_projects()


        # =====================================================
        # YEAR FILTER
        # =====================================================

        year_frame = ttk.Frame(
            project_frame
        )

        year_frame.pack(
            fill="x",
            pady=(0, 10)
        )


        ttk.Label(
            year_frame,
            text="Projektjahr:"
        ).pack(
            side="left"
        )


        years = sorted(
            {
                project["year"]
                for project in self.projects
            },
            reverse=True
        )


        self.year_combobox = ttk.Combobox(
            year_frame,
            textvariable=self.year_var,
            values=["Alle Jahre"] + years,
            state="readonly",
            width=15
        )

        self.year_combobox.pack(
            side="left",
            padx=(10, 0)
        )


        self.year_combobox.bind(
            "<<ComboboxSelected>>",
            self.update_project_list
        )


        # =====================================================
        # PROJECT SEARCH
        # =====================================================

        ttk.Label(
            project_frame,
            text="Projekt suchen:"
        ).pack(
            anchor="w"
        )


        self.search_var = tk.StringVar()


        self.search_entry = ttk.Entry(
            project_frame,
            textvariable=self.search_var
        )

        self.search_entry.pack(
            fill="x",
            pady=(5, 10)
        )


        self.search_entry.bind(
            "<KeyRelease>",
            self.update_project_list
        )


        # =====================================================
        # PROJECT LIST
        # =====================================================

        project_list_frame = ttk.Frame(
            project_frame
        )

        project_list_frame.pack(
            fill="both",
            expand=True
        )


        self.project_list = tk.Listbox(
            project_list_frame,
            height=12
        )

        self.project_list.pack(
            side="left",
            fill="both",
            expand=True
        )


        project_scrollbar = ttk.Scrollbar(
            project_list_frame,
            orient="vertical",
            command=self.project_list.yview
        )

        project_scrollbar.pack(
            side="right",
            fill="y"
        )


        self.project_list.config(
            yscrollcommand=project_scrollbar.set
        )


        self.project_list.bind(
            "<<ListboxSelect>>",
            self.on_project_selected
        )


        # =====================================================
        # SELECTED PROJECT
        # =====================================================

        self.selected_project_label = ttk.Label(
            project_frame,
            text="Kein Projekt ausgewählt."
        )

        self.selected_project_label.pack(
            anchor="w",
            pady=(10, 0)
        )


        # =====================================================
        # SAVE BUTTON
        # =====================================================

        button_frame = ttk.Frame(
            main_frame
        )

        button_frame.grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="ew",
            pady=(15, 0)
        )


        self.save_button = ttk.Button(
            button_frame,
            text="Ablegen",
            command=self.save_email_and_attachments
        )

        self.save_button.pack(
            side="right",
            ipadx=25,
            ipady=5
        )


        # =====================================================
        # INITIALIZE
        # =====================================================

        self.update_project_list()

        # Correct initial visibility of attachment destination
        self.update_attachment_destination_visibility()

    def update_attachment_list(self):
        """
        Display all real email attachments as checkboxes.
        """

        # Remove old widgets from attachment frame
        for widget in self.attachment_frame.winfo_children():
            widget.destroy()

        self.attachment_vars = []

        if not self.email:
            ttk.Label(
                self.attachment_frame,
                text="Noch keine E-Mail ausgewählt."
            ).pack(anchor="w")

            return

        attachments = self.email["attachments"]

        if not attachments:
            ttk.Label(
                self.attachment_frame,
                text="Keine Anhänge vorhanden."
            ).pack(anchor="w")

            return

        # Buttons
        button_frame = ttk.Frame(
            self.attachment_frame
        )

        button_frame.pack(
            fill="x",
            pady=(0, 5)
        )

        ttk.Button(
            button_frame,
            text="Alle auswählen",
            command=self.select_all_attachments
        ).pack(
            side="left"
        )

        ttk.Button(
            button_frame,
            text="Keine auswählen",
            command=self.deselect_all_attachments
        ).pack(
            side="left",
            padx=(5, 0)
        )

        # Attachment checkboxes
        for attachment in attachments:

            selected = tk.BooleanVar(
                value=True
            )

            self.attachment_vars.append({
                "variable": selected,
                "attachment": attachment
            })

            checkbox = ttk.Checkbutton(
                self.attachment_frame,
                text=attachment["filename"],
                variable=selected,
                command=self.update_attachment_destination_visibility
            )

            checkbox.pack(
                anchor="w",
                pady=2
            ) 
            self.update_attachment_destination_visibility()

    def select_all_attachments(self):

        for item in self.attachment_vars:
            item["variable"].set(True)

        self.update_attachment_destination_visibility()

    def deselect_all_attachments(self):

        for item in self.attachment_vars:
            item["variable"].set(False)

        self.update_attachment_destination_visibility()

    def update_name1_folders(self):
        """
        Load existing name1 folders for the selected project.
        """

        self.name1_var.set("")

        if not self.selected_project:
            self.name1_combobox["values"] = []
            return

        attachment_root = (
            Path(self.selected_project["path"])
            / self.config["paths"]["attachment_folder"]
        )

        if not attachment_root.exists():
            self.name1_combobox["values"] = []
            return

        try:

            folders = sorted(
                [
                    folder.name
                    for folder in attachment_root.iterdir()
                    if folder.is_dir()
                ],
                key=str.lower
            )

        except PermissionError:

            folders = []

        self.name1_combobox["values"] = folders

    def save_email_and_attachments(self):
        """
        Validate all required information and save
        the email and selected attachments.
        """

        # -------------------------
        # Validate email
        # -------------------------

        if self.email is None:
            messagebox.showwarning(
                "E-Mail fehlt",
                "Bitte zuerst eine E-Mail auswählen."
            )
            return


        # -------------------------
        # Validate project
        # -------------------------

        if self.selected_project is None:
            messagebox.showwarning(
                "Projekt fehlt",
                "Bitte zuerst ein Projekt auswählen."
            )
            return


        # -------------------------
        # Get selected attachments
        # -------------------------

        selected_attachments = (
            self.get_selected_attachments()
        )


        # -------------------------
        # Validate attachment data
        # only if attachments are selected
        # -------------------------

        if selected_attachments:

            name1 = self.name1_var.get().strip()
            name2 = self.name2_var.get().strip()

            if not name1:
                messagebox.showwarning(
                    "Angabe fehlt",
                    "Bitte Gemeinde, Organisation oder Firma angeben."
                )
                return

            if not name2:
                messagebox.showwarning(
                    "Angabe fehlt",
                    "Bitte die Art der Unterlagen angeben."
                )
                return


        # -------------------------
        # Save everything
        # -------------------------

        try:

            # Save email
            saved_email = save_email(
                self.email,
                self.selected_project,
                self.config,
                self.email_direction.get()
            )


            # -------------------------
            # Save attachments
            # -------------------------

            saved_files = []

            if selected_attachments:

                name1 = sanitize_filename(
                    self.name1_var.get().strip()
                )

                name2 = sanitize_filename(
                    self.name2_var.get().strip()
                )


                attachment_root = (
                    Path(self.selected_project["path"])
                    / self.config["paths"]["attachment_folder"]
                )


                name1_folder = (
                    attachment_root
                    / name1
                )


                name1_folder.mkdir(
                    parents=True,
                    exist_ok=True
                )


                attachment_destination = (
                    create_attachment_destination(
                        name1_folder=name1_folder,
                        name2=name2,
                        email_date=self.email["date"]
                    )
                )


                saved_files = save_attachments(
                    selected_attachments,
                    attachment_destination
                )


            # -------------------------
            # Success
            # -------------------------

            message = (
                "Die E-Mail wurde erfolgreich abgelegt."
            )

            if saved_files:

                message += (
                    f"\n\n{len(saved_files)} "
                    f"Anhang/Anhänge wurden gespeichert."
                )

            # Remember project
            self.add_recent_project(
                self.selected_project
            )

            messagebox.showinfo(
                "Ablage abgeschlossen",
                message
            )

            self.reset_form()


        except Exception as error:

            messagebox.showerror(
                "Fehler beim Speichern",
                f"Die Ablage konnte nicht abgeschlossen werden:\n\n{error}"
            )

    def select_email(self):

        msg_path = filedialog.askopenfilename(
            title="MSG-Datei auswählen",
            filetypes=[
                ("Outlook MSG", "*.msg"),
                ("Alle Dateien", "*.*")
            ]
        )

        if not msg_path:
            return

        self.load_email(Path(msg_path))

    def load_email(self, msg_path):

        try:
            self.email = read_email(
                Path(msg_path)
            )

        except Exception as error:
            messagebox.showerror(
                "Fehler",
                f"E-Mail konnte nicht gelesen werden:\n\n{error}"
            )

            self.email = None
            return

        self.email_subject_label.config(
            text=f"Betreff: {self.email['subject']}"
        )

        self.email_sender_label.config(
            text=f"Absender: {self.email['sender']}"
        )

        self.email_date_label.config(
            text=f"Datum: {self.email['date']}"
        )

        attachment_count = len(
            self.email["attachments"]
        )

        self.email_attachment_label.config(
            text=f"Anhänge: {attachment_count}"
        )

    
        self.update_attachment_list()

    def update_project_list(self, event=None):

        search_term = self.search_var.get().strip()
        selected_year = self.year_var.get()

        # Start with all projects
        results = self.projects

        # Filter by year
        if selected_year != "Alle Jahre":

            results = [
                project
                for project in results
                if project["year"] == selected_year
            ]

        # Filter by search term
        if search_term:

            results = search_projects(
                results,
                search_term
            )

        self.current_results = results

        self.project_list.delete(
            0,
            tk.END
        )

        for project in results:

            display_text = (
                f"{project['project_number']} | "
                f"{project['name']}"
            )

            self.project_list.insert(
                tk.END,
                display_text
            )

    def on_project_selected(self, event=None):

        selection = self.project_list.curselection()

        if not selection:
            return

        index = selection[0]

        self.selected_project = (
            self.current_results[index]
        )

        self.selected_project_label.config(
            text=(
                f"Ausgewählt: "
                f"{self.selected_project['project_number']} | "
                f"{self.selected_project['name']}"
            )
        )
        # Load existing organisations / companies
        self.update_name1_folders()

    def get_selected_attachments(self):
        """
        Return only the attachments selected by the user.
        """

        selected = []

        for item in self.attachment_vars:

            if item["variable"].get():
                selected.append(
                    item["attachment"]
                )

        return selected

    def update_attachment_destination_visibility(self):
        """
        Show attachment destination settings only
        when at least one attachment is selected.
        """

        has_selected_attachments = any(
            item["variable"].get()
            for item in self.attachment_vars
        )

        if has_selected_attachments:

            if not self.attachment_destination_frame.winfo_ismapped():
                self.attachment_destination_frame.pack(
                    fill="x",
                    pady=10
                )

        else:

            if self.attachment_destination_frame.winfo_ismapped():
                self.attachment_destination_frame.pack_forget()

    def update_recent_projects(self):
        """
        Update the buttons for the three most recently used projects.
        """

        # Remove old buttons
        for widget in self.recent_projects_frame.winfo_children():
            widget.destroy()

        if not self.recent_projects:

            ttk.Label(
                self.recent_projects_frame,
                text="Noch keine Projekte verwendet."
            ).pack(
                anchor="w"
            )

            return

        for project in self.recent_projects[:3]:

            button = ttk.Button(
                self.recent_projects_frame,
                text=(
                    f"{project['project_number']} | "
                    f"{project['name']}"
                ),
                command=lambda p=project: self.select_recent_project(p)
            )

            button.pack(
                fill="x",
                pady=2
            )

    def select_recent_project(self, project):
        """
        Select a project from the recent projects list.
        """

        self.selected_project = project

        self.selected_project_label.config(
            text=(
                f"Ausgewählt: "
                f"{project['project_number']} | "
                f"{project['name']}"
            )
        )

        self.update_name1_folders()

    def add_recent_project(self, project):
        """
        Add a project to the recent project list
        and save the list permanently.
        """

        # Remove existing occurrence
        self.recent_projects = [
            existing_project
            for existing_project in self.recent_projects
            if existing_project["path"] != project["path"]
        ]

        # Put newest project first
        self.recent_projects.insert(
            0,
            project
        )

        # Keep only the latest three
        self.recent_projects = (
            self.recent_projects[:3]
        )

        # Save project paths
        self.user_settings["recent_projects"] = [
            str(project["path"])
            for project in self.recent_projects
        ]

        save_user_settings(
            self.user_settings
        )

        self.update_recent_projects()

    def reset_form(self):
        """
        Reset the GUI after a successful filing.
        Recent projects remain unchanged.
        """

        # Email
        self.email = None

        self.email_subject_label.config(
            text="Betreff: -"
        )

        self.email_sender_label.config(
            text="Absender: -"
        )

        self.email_date_label.config(
            text="Datum: -"
        )

        self.email_attachment_label.config(
            text="Anhänge: -"
        )

        # Attachments
        self.attachment_vars = []
        self.update_attachment_list()

        # Email direction
        self.email_direction.set("Ein")

        # Project
        self.selected_project = None

        self.selected_project_label.config(
            text="Kein Projekt ausgewählt."
        )

        # Search
        self.search_var.set("")

        # Year filter
        self.year_var.set("Alle Jahre")

        # Refresh project list
        self.update_project_list()

        # Attachment destination
        self.name1_var.set("")
        self.name2_var.set("")

        # Hide attachment destination
        self.update_attachment_destination_visibility()

        # Focus ready for next email
        self.select_email_button.focus_set()

    def on_email_drop(self, event):
        """
        Handle a dropped MSG file.
        """

        dropped_files = self.root.tk.splitlist(
            event.data
        )

        if not dropped_files:
            return

        msg_path = Path(
            dropped_files[0]
        )

        if msg_path.suffix.lower() != ".msg":

            messagebox.showwarning(
                "Ungültige Datei",
                "Bitte eine Outlook-MSG-Datei verwenden."
            )

            return

        self.load_email(msg_path)

# ==========================================
# MAIN
# ==========================================



def main():
    """Main entry point of the application."""

    config = load_config()

    root_folder = Path(config["paths"]["root_folder"])
    msg_file = Path('C:/Users/jakobp/Desktop/test_mit_anhang.msg')


    print("Email Tool started")
    print(f"Root folder: {root_folder}")
    print(f"Test email:  {msg_file}")

    if not root_folder.exists():
        print("WARNING: Root folder does not exist.")

    if not msg_file.exists():
        print("ERROR: Test email does not exist.")
        return
    
    
    email = read_email(msg_file)

    print("\n--- EMAIL ---")
    print(f"Subject: {email['subject']}")
    print(f"Sender:  {email['sender']}")
    print(f"To:      {email['to']}")
    print(f"CC:      {email['cc']}")
    print(f"Date:    {email['date']}")


    print("\n--- REAL ATTACHMENTS ---")

    if email["attachments"]:

        for attachment in email["attachments"]:
            print(f"- {attachment['filename']}")
            print(f"  CID: {attachment['content_id']}")

    else:
        print("No real attachments found.")


    print("\n--- INLINE ATTACHMENTS ---")

    if email["inline_attachments"]:

        for attachment in email["inline_attachments"]:
            print(f"- {attachment['filename']}")
            print(f"  CID: {attachment['content_id']}")
            print(f"  Reason: {attachment['inline_reason']}")

    else:
        print("No inline attachments found.")



    projects = scan_projects(
    root_folder,
    config
    )
    print(f"\n{len(projects)} projects found.")

    if not projects:
        print("No projects available.")
        return

    
    # Start GUI
    root = TkinterDnD.Tk()

    app = EmailToolGUI(
        root,
        config,
        projects
    )

    root.mainloop()

if __name__ == "__main__":
    main()  


'''
    selected_project = select_project(projects)

    if selected_project is None:
        print("\nProject selection cancelled.")
        return


    print(
        f"\nSelected project: "
        f"{selected_project['project_number']} | "
        f"{selected_project['name']}"
    )


    # -------------------------
    # Select attachments
    # -------------------------

    selected_attachments = select_attachments(
        email["attachments"]
    )

    attachment_destination = None


    # -------------------------
    # Attachment destination
    # -------------------------

    if selected_attachments:

        name1_folder = select_name1_folder(
            selected_project["path"],
            config
        )

        name2 = input(
            "\nArt der Unterlagen: "
        ).strip()

        attachment_destination = (
            create_attachment_destination(
                name1_folder=name1_folder,
                name2=name2,
                email_date=None
            )
        )

        print("\nZielordner für Anhänge:")
        print(attachment_destination)


    # -------------------------
    # Save email
    # -------------------------

    email_direction = select_email_direction()

    saved_email = save_email(
        email,
        selected_project,
        config,
        email_direction
    )

    print("\nE-Mail gespeichert:")
    print(saved_email)


    # -------------------------
    # Save attachments
    # -------------------------

    if selected_attachments:

        saved_files = save_attachments(
            selected_attachments,
            attachment_destination
        )

        print("\nGespeicherte Anhänge:")

        for file in saved_files:
            print(f"- {file}")


    print("\nAblage abgeschlossen.")

if __name__ == "__main__":
    main()
'''