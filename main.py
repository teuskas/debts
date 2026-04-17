import argparse

from auth import get_dropbox_client
from desktop_app import DEFAULT_DROPBOX_PATH, run_desktop_app


def cmd_whoami() -> None:
    dbx = get_dropbox_client()
    account = dbx.users_get_current_account()
    print(f"Connesso come: {account.name.display_name} ({account.email})")


def cmd_list(path: str, limit: int) -> None:
    dbx = get_dropbox_client()
    result = dbx.files_list_folder(path=path)
    print(f"Contenuti in '{path}':")

    shown = 0
    for entry in result.entries:
        kind = "DIR" if entry.__class__.__name__.endswith("FolderMetadata") else "FILE"
        print(f"- [{kind}] {entry.name}")
        shown += 1
        if shown >= limit:
            break

    if shown == 0:
        print("(nessun elemento)")


def main() -> None:
    parser = argparse.ArgumentParser(description="Debts standalone - utility Dropbox")
    sub = parser.add_subparsers(dest="command")

    sub.add_parser("whoami", help="Mostra account Dropbox connesso")

    list_cmd = sub.add_parser("list", help="Elenca file/cartelle in un path Dropbox")
    list_cmd.add_argument("--path", default="", help="Path Dropbox da elencare (default root)")
    list_cmd.add_argument("--limit", type=int, default=20, help="Massimo elementi da stampare")

    ui_cmd = sub.add_parser("ui", help="Avvia la UI desktop standalone")
    ui_cmd.add_argument(
        "--path",
        default=DEFAULT_DROPBOX_PATH,
        help="Path Dropbox del file ODS (default /Me/DEBITI/RM+RF+RC.ods)",
    )

    args = parser.parse_args()

    if args.command in (None, "whoami"):
        cmd_whoami()
        return

    if args.command == "list":
        cmd_list(args.path, args.limit)
        return

    if args.command == "ui":
        run_desktop_app(dropbox_path=args.path)
        return

    parser.print_help()


if __name__ == "__main__":
    main()

