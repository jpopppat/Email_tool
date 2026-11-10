from pathlib import Path
import extract_msg
import yaml
import re
from datetime import datetime
import shutil

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

def scan_projects(root_folder):
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
        r"^(?P<year>\d{2})_(?P<number>\d{3})_(?P<name>.+)$"
    )

    # Scan root folder
    try:
        year_folders = list(root_folder.iterdir())

    except PermissionError:
        raise PermissionError(
            f"No permission to access root folder: {root_folder}"
        )

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
    
def main():
    """Main entry point of the application."""

    config = load_config()

    root_folder = Path(config["paths"]["root_folder"])
    msg_file = Path('C:/Users/jakobp/Desktop/Fw_ Vorbereitung Dienstbarkeitsverträge Arzler Alm-Leitung.msg')

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

    projects = scan_projects(root_folder)

    print(f"\n{len(projects)} projects found.")

    if not projects:
        print("No projects available.")
        return


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