import requests

 # Check for README function: searches entire repo tree for any README file (case-insensitive, can be in subdirectories)
def get_readme_info(owner, repo, headers=None):
    """Search entire repo tree for any README file."""
    # Get default branch's tree recursively
    tree_url = f"https://api.github.com/repos/{owner}/{repo}/git/trees/HEAD?recursive=1"
    response = requests.get(tree_url, headers=headers)
    
    if response.status_code != 200:
        return "Unknown"
    
    tree = response.json().get("tree", [])
    
    readme_files = [
        item["path"] for item in tree
        if item["type"] == "blob" and 
        item["path"].split("/")[-1].lower().startswith("readme")
    ]
    
    if not readme_files:
        return "No"
    
    return f"Yes ({len(readme_files)}): " + ", ".join(readme_files)
