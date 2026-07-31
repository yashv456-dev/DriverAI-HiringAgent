"""Heal all Resume URL, Resume Folder Path, and Resume Link cells across both
main (HiringAgent_P1_Candidates) and rejected (RejectedCandidates) tables on SharePoint
to reflect the new folder name 'Candidate_Resumes'.
"""

import sys
import hiring_agent.config
from hiring_agent.sharepoint_scoring import _RESUME_LINK_FORMULA
from sharepoint_client import SharePointClient

def heal_all_links():
    client = SharePointClient()
    print("=" * 60)
    print("HEALING RESUME LINKS & FOLDER PATHS ON SHAREPOINT")
    print("=" * 60)
    
    # 1. Main Sheet
    main_cols = client.table_columns()
    main_rows = client.list_rows()
    print(f"\n1. Auditing {len(main_rows)} rows on Main Sheet ('HiringAgent_P1_Candidates')...")
    
    main_fixed = 0
    for r in main_rows:
        idx = r["index"]
        vals = r["values"]
        old_url = str(vals.get("Resume URL", ""))
        old_folder = str(vals.get("Resume Folder Path", ""))
        
        needs_update = False
        updates = {}
        
        if "Downloaded_Resumes" in old_url:
            new_url = old_url.replace("Downloaded_Resumes", "Candidate_Resumes")
            updates["Resume URL"] = new_url
            needs_update = True
            
        if "Downloaded_Resumes" in old_folder:
            new_folder = old_folder.replace("Downloaded_Resumes", "Candidate_Resumes")
            updates["Resume Folder Path"] = new_folder
            needs_update = True
            
        if needs_update:
            # Rebuild formula if Resume URL present
            url_to_use = updates.get("Resume URL", old_url)
            if url_to_use:
                updates["Resume Link"] = _RESUME_LINK_FORMULA
            client.update_row(idx, updates)
            main_fixed += 1

    print(f"   Done! Updated {main_fixed} main sheet rows.")

    # 2. Rejected Sheet
    rej_cols = client._table_columns_of("RejectedCandidates")
    raw_rej = client._rows_paged(client._wb_base() + "/tables/RejectedCandidates/rows")
    rej_rows = []
    for r in raw_rej:
        vals = r.get("values", [[]])[0]
        mapped = {rej_cols[i]: (vals[i] if i < len(vals) else "") for i in range(len(rej_cols))}
        if any(str(v or "").strip() for v in mapped.values()):
            rej_rows.append({"index": r.get("index"), "values": mapped})

    print(f"\n2. Auditing {len(rej_rows)} rows on Rejected Sheet ('RejectedCandidates')...")
    
    rej_fixed = 0
    for r in rej_rows:
        idx = r["index"]
        vals = r["values"]
        old_url = str(vals.get("Resume URL", ""))
        old_folder = str(vals.get("Resume Folder Path", ""))
        
        needs_update = False
        updates = {}
        
        if "Downloaded_Resumes" in old_url:
            new_url = old_url.replace("Downloaded_Resumes", "Candidate_Resumes")
            updates["Resume URL"] = new_url
            needs_update = True
            
        if "Downloaded_Resumes" in old_folder:
            new_folder = old_folder.replace("Downloaded_Resumes", "Candidate_Resumes")
            updates["Resume Folder Path"] = new_folder
            needs_update = True
            
        if needs_update:
            url_to_use = updates.get("Resume URL", old_url)
            if url_to_use:
                updates["Resume Link"] = _RESUME_LINK_FORMULA
            client.update_rejected_row(idx, updates)
            rej_fixed += 1

    print(f"   Done! Updated {rej_fixed} rejected sheet rows.")
    print("=" * 60)
    print("LINK HEALING COMPLETE: All URLs now point to /Candidate_Resumes/")

if __name__ == "__main__":
    heal_all_links()
