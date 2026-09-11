"""Fix all Resume URL, Resume Folder Path, and Resume Link entries in the
RejectedCandidates table on SharePoint so they point to /Candidate_Resumes/.
"""

import hiring_agent.config
from hiring_agent.sharepoint_scoring import _RESUME_LINK_FORMULA
from sharepoint_client import SharePointClient

def fix_rejected_sheet():
    client = SharePointClient()
    table = "RejectedCandidates"
    cols = client._table_columns_of(table)
    raw = client._rows_paged(f"{client._wb_base()}/tables/{table}/rows")
    
    print(f"Auditing {len(raw)} rows in table '{table}'...")
    fixed = 0
    
    for r in raw:
        idx = r.get("index")
        vals = r["values"][0] if r.get("values") else []
        mapped = {cols[i]: (vals[i] if i < len(vals) else "") for i in range(len(cols))}
        
        old_url = str(mapped.get("Resume URL", ""))
        old_folder = str(mapped.get("Resume Folder Path", ""))
        
        updates = {}
        if "Downloaded_Resumes" in old_url:
            updates["Resume URL"] = old_url.replace("Downloaded_Resumes", "Candidate_Resumes")
        if "Downloaded_Resumes" in old_folder:
            updates["Resume Folder Path"] = old_folder.replace("Downloaded_Resumes", "Candidate_Resumes")
            
        if updates:
            merged = dict(mapped)
            merged.update(updates)
            
            # Preserve hyperlink formula in Resume Link if present
            if "Resume Link" in cols and merged.get("Resume URL"):
                merged["Resume Link"] = _RESUME_LINK_FORMULA
                
            new_row_values = [[merged.get(c, "") for c in cols]]
            url = f"{client._wb_base()}/tables/{table}/rows/itemAt(index={idx})"
            client._req("PATCH", url, json={"values": new_row_values})
            fixed += 1
            
    print(f"Successfully fixed {fixed} rows on the Rejected sheet!")

if __name__ == "__main__":
    fix_rejected_sheet()
