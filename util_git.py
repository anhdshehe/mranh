"""
Git util
"""

from subprocess import check_output
from typing import List


def force_push(remote_name, revision, file_names: List[str]):
    """
    force push files in remote

    Args:
        remote_name (_type_): _description_
        revision (_type_): _description_
        file_names (List[str]): _description_
    """
    # Push file to git
    check_output(f"git pull {remote_name} {revision}", shell=True).decode()
    file_names = ' '.join(file_names)
    check_output(f"git add {file_names}", shell=True).decode()
    check_output(f"git commit -m \"Add {file_names}\"", shell=True).decode()
    check_output(f"git push {remote_name} {revision}", shell=True).decode()

    return None
