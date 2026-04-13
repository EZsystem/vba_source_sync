# -*- coding: utf-8 -*-
import os, json

def force_write():
    path_ab = r'D:\My_code\11_workspaces\VBA_manager\vba_source_sync\workspace\data_ab\table_values\tbl_内訳.md'
    path_om = r'D:\My_code\11_workspaces\VBA_manager\vba_source_sync\workspace\data_oomati\table_values\tbl_内訳.md'
    dest = r'D:\My_code\11_workspaces\VBA_manager\vba_source_sync\workspace\data_oomati\BELCA_Logic_Master.md'
    
    master_data = []
    # AB
    if os.path.exists(path_ab):
        with open(path_ab, 'r', encoding='utf-8', errors='ignore') as f:
            cur_a = None
            for line in f:
                if not line.strip().startswith('|') or '---' in line or 'O-ID' in line: continue
                c = [x.strip() for x in line.split('|')]
                if len(c) < 17: continue
                if c[2] == 'a': cur_a = c
                elif c[2] == 'b' and cur_a:
                    if c[16] and 'VECD' not in line:
                        master_data.append({'category': c[6], 'name': f'{cur_a[9]}::{c[9]}', 'spec': f'{cur_a[10]}::{c[10]}', 'belca': c[16]})
                    cur_a = None
    # Oomati
    if os.path.exists(path_om):
        with open(path_om, 'r', encoding='utf-8', errors='ignore') as f:
            for line in f:
                if not line.strip().startswith('|') or '---' in line or 'O-ID' in line: continue
                c = [x.strip() for x in line.split('|')]
                if len(c) < 17 or not c[10]: continue
                if c[16] and 'VECD' not in line:
                    master_data.append({'category': c[6], 'name': c[10], 'spec': c[11], 'belca': c[16]})

    unique = []
    seen = set()
    for x in master_data:
        key = (x['name'], x['spec'], x['belca'])
        if key not in seen:
            unique.append(x)
            seen.add(key)

    # WRITE DIRECTLY
    with open(dest, 'w', encoding='utf-8') as f:
        f.write('# BELCA判定ルール・マスター定義\n\n')
        f.write('## 3. AI学習・参照用データコア (Structured JSON)\n\n')
        f.write('```json\n[\n')
        for i, item in enumerate(unique):
            suffix = ',' if i < len(unique) - 1 else ''
            f.write('  ' + json.dumps(item, ensure_ascii=False) + suffix + '\n')
        f.write(']\n```\n')
    
    print(f'PHYSICAL_FILE_SAVED: {len(unique)} rules recorded.')

if __name__ == "__main__":
    force_write()
