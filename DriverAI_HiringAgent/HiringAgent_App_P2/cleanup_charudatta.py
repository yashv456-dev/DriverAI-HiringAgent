import hiring_agent.config
from sharepoint_client import SharePointClient

client = SharePointClient()
main_rows = client.list_rows()

char_main_indices = [r["index"] for r in main_rows if "Charudatta" in str(r["values"].get("Full Name"))]

print(f"Found {len(char_main_indices)} Charudatta rows on Main sheet. Cleaning up...")
for idx in sorted(char_main_indices, reverse=True):
    client.delete_row(idx)
    print(f"Deleted main sheet row index {idx}")

print("Charudatta cleanup complete!")
