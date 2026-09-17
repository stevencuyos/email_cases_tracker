import sys

def replace_in_file(filepath, search_str, replace_str):
    with open(filepath, 'r') as f:
        content = f.read()

    if search_str in content:
        new_content = content.replace(search_str, replace_str)
        with open(filepath, 'w') as f:
            f.write(new_content)
        print("Success")
    else:
        print("Search string not found")

