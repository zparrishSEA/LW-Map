import pandas as pd
import numpy as np
import json

print("1. Loading Master Databases...")
state_df = pd.read_excel('State Data.xlsx', dtype=str).fillna('')
master_df = pd.read_excel('Master_Mapper_Updated.xlsx', dtype=str).fillna('')

# Ensure all required columns exist to prevent crashes
required_cols = [
    'Row Type', 'Regions in Dist.', 'district_id', 'school_id',
    'School Search Name', 'School Name', 'Program Launch Step', 'Signatures', 'LW Dist. URL', 'State', 'Dist/Region Name',
    'Dist. Type', 'Dist. Charter Status', 'Mailing Address', 'LW Director',
    'LW Director Email', 'School Type', 'Dist. Name',
    'RTRI Policy', 'Public Records', 'Local Issues', 'Local Media', 'Connected Churches', 'Local Business Sponsors'
	'Weekly Change', 'Sig. Change', 'Step Change', 'Old Signatures', 'Old Step'
]
for col in required_cols:
    if col not in master_df.columns:
        master_df[col] = ""

master_df['Dist/Region Name'] = master_df['Dist/Region Name'].astype(str).str.strip()
master_df['Dist. Code'] = master_df['Dist. Code'].apply(lambda x: str(x).replace('.0', '').strip().zfill(7) if str(x).strip() not in ['nan', '', 'None'] else "")

print("2. Parsing Sequential Data Groups...")
# Iterate sequentially: district row followed by ALL its respective school rows
groups = []
current_group = None

for _, row in master_df.iterrows():
    row_type = str(row.get('Row Type', '')).strip().lower()

    if row_type == 'district':
        if current_group is not None:
            groups.append(current_group)
        current_group = {'meta': row, 'schools': []}

    elif row_type == 'school':
        if current_group is not None:
            current_group['schools'].append(row)

if current_group is not None:
    groups.append(current_group)

print("3. Extracting State Database Info...")
us_state_abbrev = {
    'AL': 'Alabama', 'AK': 'Alaska', 'AZ': 'Arizona', 'AR': 'Arkansas', 'CA': 'California', 'CO': 'Colorado',
    'CT': 'Connecticut', 'DE': 'Delaware', 'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii', 'ID': 'Idaho',
    'IL': 'Illinois', 'IN': 'Indiana', 'IA': 'Iowa', 'KS': 'Kansas', 'KY': 'Kentucky', 'LA': 'Louisiana',
    'ME': 'Maine', 'MD': 'Maryland', 'MA': 'Massachusetts', 'MI': 'Michigan', 'MN': 'Minnesota', 'MS': 'Mississippi',
    'MO': 'Missouri', 'MT': 'Montana', 'NE': 'Nebraska', 'NV': 'Nevada', 'NH': 'New Hampshire', 'NJ': 'New Jersey',
    'NM': 'New Mexico', 'NY': 'New York', 'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio', 'OK': 'Oklahoma',
    'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island', 'SC': 'South Carolina', 'SD': 'South Dakota',
    'TN': 'Tennessee', 'TX': 'Texas', 'UT': 'Utah', 'VT': 'Vermont', 'VA': 'Virginia', 'WA': 'Washington',
    'WV': 'West Virginia', 'WI': 'Wisconsin', 'WY': 'Wyoming', 'DC': 'District of Columbia'
}

state_dict = {}
for _, row in state_df.iterrows():
    s_name = str(row.get('State', '')).strip()
    if len(s_name) == 2:
        s_name_upper = s_name.upper()
        if s_name_upper in us_state_abbrev:
            s_name = us_state_abbrev[s_name_upper]

    if s_name:
        state_dict[s_name] = {k: str(v).strip() for k, v in row.items() if str(v).strip() != '' and str(v).strip().lower() != 'nan'}

print("4. Building the Master JSON Architecture...")
output_data = []

def safe_int(val):
    res = pd.to_numeric(val, errors='coerce')
    return int(res) if pd.notna(res) else 0

def get_exact_school_name(s_row):
    val = str(s_row.get('School Search Name', '')).strip()
    if not val:
        val = str(s_row.get('School Name', '')).strip()
    return val

def build_school_obj(s_row, phase="0", grades=""):
    return {
        "School Name": get_exact_school_name(s_row),
        "School LAT": s_row.get('School LAT', ''),
        "School LON": s_row.get('School LON', ''),
        "School Address": s_row.get('School Address', ''),
        "School Phone": s_row.get('School Phone', ''),
        "School Type": s_row.get('School Type', ''),
        "School Total Students": s_row.get('School Total Students', ''),
        "Charter School": s_row.get('Charter School', ''),
        "School Student/Teacher Ratio": s_row.get('School Student/Teacher Ratio', ''),
        "School Lunch Program Status": s_row.get('School Lunch Program Status', ''),
        "School Free and Reduced Lunch Student Totals": s_row.get('School Free and Reduced Lunch Student Totals', ''),
        "Grades Served": grades,
        "School URL": s_row.get('School URL', ''),
        "Program Launch Step": str(phase),
        "Director Name": str(s_row.get('LW Director', '')).strip(),
        "Director Email": str(s_row.get('LW Director Email', '')).strip(),
        "Mailing Address": str(s_row.get('Mailing Address', '')).strip(),
        "LW URL": str(s_row.get('LW Dist. URL', '')).strip(),
        "Connected Churches": str(s_row.get('Connected Churches', '')).strip(),
        "Local Business Sponsors": str(s_row.get('Local Business Sponsors', '')).strip(),
        "Weekly Change": str(s_row.get('Weekly Change', '')).strip(),
		"Signatures": str(s_row.get('Signatures', '')).strip(),
        "Sig. Change": str(s_row.get('Sig. Change', '')).strip(),
        "Step Change": str(s_row.get('Step Change', '')).strip(),
        "Old Signatures": str(s_row.get('Old Signatures', '')).strip(),
        "Old Step": str(s_row.get('Old Step', '')).strip()
    }

def is_non_traditional(df_rows):
    if df_rows.empty: return False
    for _, s in df_rows.iterrows():
        s_type = str(s.get('School Type', '')).strip()
        if s_type.startswith('1') or s_type == '':
            return False
    return True

# --- PROCESS SEQUENTIAL GROUPS ---
for grp in groups:
    r_meta = grp['meta']
    schools_df = pd.DataFrame(grp['schools'])

    if not schools_df.empty:
        schools_df['Dedup_Name'] = schools_df.apply(get_exact_school_name, axis=1)
        schools_df = schools_df.drop_duplicates(subset=['Dedup_Name'])

    dist_code = str(r_meta.get('Dist. Code', '')).replace('.0', '').strip().zfill(7)
    dist_name = str(r_meta.get('Dist. Name', '')).strip()
    if not dist_code: dist_code = "UNKNOWN_" + dist_name.replace(" ", "")[:8]

    regions_count = safe_int(r_meta.get('Regions in Dist.', 1))
    is_umbrella = regions_count > 1

    dist_type_code = str(r_meta.get('Dist. Type', '')).strip()

    base_meta = {
        "Dist. Code": dist_code, "Dist. Name": dist_name, "State": r_meta.get('State', ''),
        "Dist. Address": r_meta.get('Dist. Address', ''), "Dist. Phone": r_meta.get('Dist. Phone', ''),
        "Dist. URL": r_meta.get('Dist. URL', ''), "Dist. Type": r_meta.get('Dist. Type', ''),
        "Dist. Locale": r_meta.get('Dist. Locale', ''), "Dist. Charter Status": r_meta.get('Dist. Charter Status', ''),
        "Dist. Total Number of Schools": r_meta.get('Dist. Total Number of Schools', ''),
        "Dist. Total Number of Charter Schools": str(r_meta.get('Dist. Total Number of Charter Schools', '')),
        "Dist. Total Students, All Grades": r_meta.get('Dist. Total Students, All Grades', ''),
        "Dist. Pupil/Teacher Ratio": r_meta.get('Dist. Pupil/Teacher Ratio', ''),
        "LW URL": r_meta.get('LW Dist. URL', ''), "Director Name": r_meta.get('LW Director', ''),
        "Director Email": r_meta.get('LW Director Email', ''), "Mailing Address": r_meta.get('Mailing Address', ''),
        "RTRI Policy": r_meta.get('RTRI Policy', ''), "Public Records": r_meta.get('Public Records', ''),
		"Local Issues": r_meta.get('Local Issues', ''),   # <-- Added back!
        "Local Media": r_meta.get('Local Media', ''),     # <-- Added back!
        "Connected Churches": r_meta.get('Connected Churches', ''),
        "Local Business Sponsors": r_meta.get('Local Business Sponsors', ''),
        "Weekly Change": r_meta.get('Weekly Change', ''),
        "Sig. Change": r_meta.get('Sig. Change', ''),
        "Step Change": r_meta.get('Step Change', ''),
        "Old Signatures": r_meta.get('Old Signatures', ''),
        "Old Step": r_meta.get('Old Step', '')
    }

    if is_umbrella and not schools_df.empty:
        # --- 1. GENERATE UNIQUE REGIONS ---
        # Ensure blank regions don't crash the script; group them safely
        schools_df['Dist/Region Name'] = schools_df['Dist/Region Name'].fillna('').astype(str).str.strip()
        schools_df['Dist/Region Name'] = schools_df['Dist/Region Name'].replace('', 'Local Program')

        unique_regions = schools_df['Dist/Region Name'].unique()

        for region_name in unique_regions:
            reg_schools = schools_df[schools_df['Dist/Region Name'] == region_name]

            reg_phase = safe_int(reg_schools['Program Launch Step'].max())
            reg_sigs = safe_int(reg_schools['Signatures'].max())
            reg_non_trad = is_non_traditional(reg_schools) or dist_type_code.startswith(('7', '8'))

            # Create a unique ID for each region
            safe_reg_id = "".join(c for c in region_name if c.isalnum())
            reg_id = f"{dist_code}_{safe_reg_id}"

            reg_obj = {
                "id": reg_id, "type": "region", "isNonTraditional": reg_non_trad,
                "phaseNum": reg_phase, "signatures": reg_sigs, "name": dist_name, "regionName": region_name,
                "meta": base_meta, "schools": []
            }

            lat_list, lon_list = [], []
            for _, s in reg_schools.iterrows():
                s_phase = safe_int(s.get('Program Launch Step', 0))
                grades = str(s.get('Grades Served', '')).replace('nan', '').strip()
                if grades != "":
                    s_lat = pd.to_numeric(s.get('School LAT'), errors='coerce')
                    s_lon = pd.to_numeric(s.get('School LON'), errors='coerce')
                    if pd.notna(s_lat) and pd.notna(s_lon):
                        lat_list.append(s_lat)
                        lon_list.append(s_lon)
                reg_obj['schools'].append(build_school_obj(s, s_phase, grades))

            # Regions calculate their coordinates based purely on their specific schools
            if len(lat_list) > 0:
                reg_obj['lat'] = sum(lat_list) / len(lat_list)
                reg_obj['lon'] = sum(lon_list) / len(lon_list)
            else:
                reg_obj['lat'] = float(r_meta['Dist. LAT']) if r_meta.get('Dist. LAT') else np.nan
                reg_obj['lon'] = float(r_meta['Dist LON']) if r_meta.get('Dist LON') else np.nan
                if pd.notna(reg_obj['lat']):
                    reg_obj['lat'] += np.random.uniform(-0.005, 0.005)
                    reg_obj['lon'] += np.random.uniform(-0.005, 0.005)

            output_data.append(reg_obj)

        # --- 2. GENERATE MASTER UMBRELLA ---
        umb_phase = safe_int(schools_df['Program Launch Step'].max())
        umb_sigs = safe_int(schools_df['Signatures'].max())
        umb_non_trad = is_non_traditional(schools_df) or dist_type_code.startswith(('7', '8'))

        umb_obj = {
            "id": dist_code,
            "type": "umbrella",
            "isNonTraditional": umb_non_trad,
            "phaseNum": umb_phase,
            "signatures": umb_sigs,
            "name": dist_name,
            "regionName": "",
            "meta": base_meta,
            "schools": [],
            # Umbrellas explicitly use official District Coordinates
            "lat": float(r_meta['Dist. LAT']) if r_meta.get('Dist. LAT') else np.nan,
            "lon": float(r_meta['Dist LON']) if r_meta.get('Dist LON') else np.nan
        }

        for _, s in schools_df.iterrows():
            s_phase = safe_int(s.get('Program Launch Step', 0))
            grades = str(s.get('Grades Served', '')).replace('nan', '').strip()
            umb_obj['schools'].append(build_school_obj(s, s_phase, grades))

        output_data.append(umb_obj)

    else:
        # --- 3. GENERATE SINGLE DISTRICT (Regions <= 1) ---
        if not schools_df.empty:
            phase = safe_int(schools_df['Program Launch Step'].max())
            sigs = safe_int(schools_df['Signatures'].max())
            non_trad = is_non_traditional(schools_df) or dist_type_code.startswith(('7', '8'))
        else:
            phase = safe_int(r_meta.get('Program Launch Step', 0))
            sigs = safe_int(r_meta.get('Signatures', 0))
            non_trad = dist_type_code.startswith(('7', '8'))

        dist_id = str(r_meta.get('district_id', '')).strip()
        obj_id = dist_id if dist_id else dist_code

        single_obj = {
            "id": obj_id, "type": "single", "isNonTraditional": non_trad,
            "phaseNum": phase, "signatures": sigs, "name": dist_name, "regionName": "",
            "meta": base_meta, "schools": [],
            "lat": float(r_meta['Dist. LAT']) if r_meta.get('Dist. LAT') else np.nan,
            "lon": float(r_meta['Dist LON']) if r_meta.get('Dist LON') else np.nan
        }

        if not schools_df.empty:
            for _, s in schools_df.iterrows():
                s_phase = safe_int(s.get('Program Launch Step', 0))
                grades = str(s.get('Grades Served', '')).replace('nan', '').strip()
                single_obj['schools'].append(build_school_obj(s, s_phase, grades))

        output_data.append(single_obj)

print("5. Saving Data by State and Generating Master Index...")
import os
if not os.path.exists('state_data'):
    os.makedirs('state_data')

master_districts = []
state_files = {}

for dist in output_data:
    # 1. Grab the state name safely
    state_name = str(dist.get('meta', {}).get('State', 'Unknown')).strip()
    if not state_name: state_name = "Unknown"

    # 2. Store the FULL, heavy district object in the state dictionary
    if state_name not in state_files:
        state_files[state_name] = []
    state_files[state_name].append(dist)

    # 3. Build a LIGHTWEIGHT version for the Master Index (just for map dots & search)
    light_dist = {
        "id": dist["id"],
        "type": dist["type"],
        "isNonTraditional": dist["isNonTraditional"],
        "phaseNum": dist["phaseNum"],
        "signatures": dist["signatures"],
        "name": dist["name"],
        "regionName": dist["regionName"],
        "lat": dist["lat"],
        "lon": dist["lon"],
        "state": state_name,
        "meta": {
            "Dist. Code": dist["meta"].get("Dist. Code", ""),
            "Weekly Change": dist["meta"].get("Weekly Change", ""),
            "Step Change": dist["meta"].get("Step Change", "")
        },
        "schools": []
    }

    # Strip the heavy text out of the schools array for the master index
    for s in dist["schools"]:
        light_dist["schools"].append({
            "School Name": s.get("School Name", ""),
            "School LAT": s.get("School LAT", ""),
            "School LON": s.get("School LON", ""),
            "Grades Served": s.get("Grades Served", ""),
            "Program Launch Step": s.get("Program Launch Step", ""),
            "Charter School": s.get("Charter School", "")
        })

    master_districts.append(light_dist)

def clean_nans(obj):
    if isinstance(obj, dict): return {k: clean_nans(v) for k, v in obj.items()}
    elif isinstance(obj, list): return [clean_nans(v) for v in obj]
    elif isinstance(obj, float) and np.isnan(obj): return ""
    return obj

# Save Master Index
final_master = { "districts": master_districts, "state_meta": state_dict }
with open('master_index.json', 'w') as f:
    json.dump(clean_nans(final_master), f, separators=(',', ':'))

# Save Individual State Files
for st, dists in state_files.items():
    safe_st = "".join([c for c in st if c.isalnum() or c == ' ']).replace(' ', '_')
    with open(f'state_data/{safe_st}.json', 'w') as f:
        json.dump(clean_nans(dists), f, separators=(',', ':'))

print(f"Complete! master_index.json and {len(state_files)} state files built.")