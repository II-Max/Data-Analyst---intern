import pandas as pd
import uuid
import warnings
import os

warnings.filterwarnings('ignore')

def safe_int_convert(val):
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return None

def standardize_date(val):
    if pd.isna(val) or str(val).strip() in ['', 'nan']: return ''
    val_str = str(val).strip()
    if val_str.endswith('--'): return val_str[:-2] + '----'
    try:
        dt = pd.to_datetime(val_str, dayfirst=True, errors='coerce')
        return dt.strftime('%d/%m/%Y') if pd.notna(dt) else val_str
    except:
        return val_str

def get_ordinal_rank(row, df):
    ten_bo = str(row.get('Tên bố', '')).strip()
    ten_me = str(row.get('Tên mẹ', '')).strip()
    gioi_tinh = str(row.get('Giới tính', '')).strip()
    
    if (not ten_bo or ten_bo.lower() == 'nan') and (not ten_me or ten_me.lower() == 'nan'): return ""
    
    siblings = df[(df['Tên bố'].str.strip() == ten_bo) & (df['Tên mẹ'].str.strip() == ten_me)]
    siblings = siblings.sort_values(by='STT')
    same_gender = siblings[siblings['Giới tính'].str.strip() == gioi_tinh]
    
    try:
        name_list = list(same_gender['Tên'].str.strip())
        rank_idx = name_list.index(str(row['Tên']).strip()) + 1
        if gioi_tinh.upper() == 'NAM':
            return "trưởng nam" if rank_idx == 1 else f"con thứ {rank_idx}"
        else:
            return "trưởng nữ" if rank_idx == 1 else f"con thứ {rank_idx}"
    except ValueError:
        return ""

def format_display_name(row):
    ten = str(row.get('Tên', '')).strip()
    me = str(row.get('Tên mẹ', '')).strip()
    bo = str(row.get('Tên bố', '')).strip()
    if me and me.lower() != 'nan': return f"{ten} (Con Mẹ {me})"
    if bo and bo.lower() != 'nan': return f"{ten} (Con Bố {bo})"
    return ten

def build_full_description(row, df):
    ten = str(row.get('Tên', '')).strip()
    if not ten or ten.lower() == 'nan': return ""

    doi_val = safe_int_convert(row.get('Đời'))
    gt = str(row.get('Giới tính', '')).strip().upper()
    
    noi_mat = str(row.get('Nơi mất', '')).strip()
    if not noi_mat or noi_mat.lower() in ['nan', '']:
        noi_mat = "Thôn Lạc Thành Nam, xã Tây Ninh, huyện Tiền Hải, tỉnh Thái Bình"

    nm_al = str(row.get('Ngày mất (AL)', '')).strip()
    
    header = f"{ten}"
    if nm_al and nm_al.lower() != 'nan': header += f", mất {nm_al} (ÂL)"
    header += f", mộ mai táng tại {noi_mat}"
    desc = header
    
    if doi_val and doi_val > 1:
        rank = get_ordinal_rank(row, df)
        bo = str(row.get('Tên bố', '')).strip()
        me = str(row.get('Tên mẹ', '')).strip()
        if bo and bo.lower() != 'nan' and me and me.lower() != 'nan':
            desc += f", là {rank} của thủy tổ ông {bo} và thủy tổ bà {me}"

    extra = ""
    if gt == 'NAM':
        wives = df[df['Tên bố'].str.strip() == ten]['Tên mẹ'].unique()
        wives = [w for w in wives if pd.notna(w) and str(w).strip() != '']
        
        if doi_val == 1:
            if len(wives) > 0:
                w_name = str(wives[0]).strip()
                children = df[(df['Tên bố'].str.strip() == ten) & (df['Tên mẹ'].str.strip() == w_name)]['Tên hiển thị'].tolist()
                extra += f". Vợ của {ten} là {w_name}, mộ mai táng tại {noi_mat}"
                if children:
                    extra += ", Cụ Tổ có các con là:\n" + "\n".join([f"{i+1}. {c}" for i, c in enumerate(children)])
        else:
            if len(wives) > 1: extra += f". Cụ ông có {len(wives)} người vợ."
            for w_name in wives:
                w_name = str(w_name).strip()
                children = df[(df['Tên bố'].str.strip() == ten) & (df['Tên mẹ'].str.strip() == w_name)]['Tên hiển thị'].tolist()
                pfx = f"\n {w_name}, mộ mai táng tại {noi_mat}" if len(wives) > 1 else f". Vợ của {ten} là {w_name}, mộ mai táng tại {noi_mat}"
                if children:
                    pfx += ", hai cụ có các con là:\n" + "\n".join([f"{i+1}. {c}" for i, c in enumerate(children)])
                extra += pfx
                
    elif gt == 'NỮ':
        husbands = df[df['Tên mẹ'].str.strip() == ten]['Tên bố'].unique()
        husbands = [h for h in husbands if pd.notna(h) and str(h).strip() != '']
        
        if doi_val == 1:
            if len(husbands) > 0:
                h_name = str(husbands[0]).strip()
                children = df[(df['Tên bố'].str.strip() == h_name) & (df['Tên mẹ'].str.strip() == ten)]['Tên hiển thị'].tolist()
                h_row = df[df['Tên'].str.strip() == h_name]
                h_mat = ""
                if not h_row.empty:
                    h_nm_al = str(h_row.iloc[0].get('Ngày mất (AL)', '')).strip()
                    if h_nm_al and h_nm_al.lower() != 'nan': h_mat = f", mất {h_nm_al} (ÂL)"
                extra += f". Chồng của {ten} là {h_name}{h_mat}, mộ mai táng tại {noi_mat}"
                if children:
                    extra += ", Thủy tổ bà có các con là:\n" + "\n".join([f"{i+1}. {c}" for i, c in enumerate(children)])
        else:
            for h_name in husbands:
                h_name = str(h_name).strip()
                children = df[(df['Tên bố'].str.strip() == h_name) & (df['Tên mẹ'].str.strip() == ten)]['Tên hiển thị'].tolist()
                pfx = f". Chồng của {ten} là {h_name}, mộ mai táng tại {noi_mat}"
                if children: pfx += ", hai cụ có các con là:\n" + "\n".join([f"{i+1}. {c}" for i, c in enumerate(children)])
                extra += pfx

    full_text = (desc + extra).replace("..", ".").replace(",.", ".").strip()
    if full_text and not full_text.endswith('.'): full_text += "."
    return full_text

input_file = 'Input_template.xlsx'
if not os.path.exists(input_file):
    if os.path.exists(f"Data/{input_file}"): input_file = f"Data/{input_file}"
    else: raise FileNotFoundError(f"Không tìm thấy file '{input_file}'.")

df = pd.read_excel(input_file, skiprows=6)
df.columns = [str(col).strip() for col in df.columns]

if 'Ngày  mất (DL)' in df.columns: df = df.rename(columns={'Ngày  mất (DL)': 'Ngày mất (DL)'})
if 'Ngày  mất (AL)' in df.columns: df = df.rename(columns={'Ngày  mất (AL)': 'Ngày mất (AL)'})
if 'Tên ' in df.columns: df = df.rename(columns={'Tên ': 'Tên'})

df = df[df['Tên'].notna()]
df = df[df['Tên'].astype(str).str.strip() != '']

df['Tên hiển thị'] = df.apply(format_display_name, axis=1)
df = df.sort_values('STT')
df['Nội tộc'], df['Thứ bậc'], df['Mã dòng họ'] = '', '', ''

for doi in df['Đời'].dropna().unique():
    doi_val = safe_int_convert(doi)
    if not doi_val: continue
    th_counter = 0
    last_th = ""
    
    for idx in df[df['Đời'] == doi].index:
        bo, me = str(df.at[idx, 'Tên bố']).strip(), str(df.at[idx, 'Tên mẹ']).strip()
        gt = str(df.at[idx, 'Giới tính']).strip().upper()
        
        is_bloodline = False
        if doi_val == 1 and gt == 'NAM': is_bloodline = True
        elif (bo != 'nan' and bo != '') or (me != 'nan' and me != ''): is_bloodline = True
        
        if is_bloodline:
            df.at[idx, 'Nội tộc'] = 'X'
            if doi_val > 1:
                th_counter += 1
                last_th = str(th_counter)
                df.at[idx, 'Thứ bậc'] = last_th
                df.at[idx, 'Mã dòng họ'] = f"Đ{doi_val}-TH{last_th}"
            else: df.at[idx, 'Mã dòng họ'] = f"Đ1"
        else:
            if doi_val > 1 and last_th != "":
                df.at[idx, 'Thứ bậc'] = last_th
                df.at[idx, 'Mã dòng họ'] = f"Đ{doi_val}-TH{last_th}"
            elif doi_val == 1: df.at[idx, 'Mã dòng họ'] = f"Đ1"

df['Mã thành viên'] = [str(uuid.uuid4()) for _ in range(len(df))]
df['key'] = df.apply(lambda r: f"{str(r.get('Tên', '')).strip()}|{safe_int_convert(r.get('Đời')) or ''}", axis=1)
name_to_uuid = dict(zip(df['key'], df['Mã thành viên']))
uuid_to_raw_name = dict(zip(df['Mã thành viên'], df['Tên']))

spouse_map = {} 
for _, row in df.iterrows():
    bo, me, doi = str(row.get('Tên bố','')).strip(), str(row.get('Tên mẹ','')).strip(), safe_int_convert(row.get('Đời'))
    if bo and me and doi and bo.lower() != 'nan' and me.lower() != 'nan':
        bo_id = name_to_uuid.get(f"{bo}|{doi-1}")
        me_id = name_to_uuid.get(f"{me}|{doi-1}")
        if bo_id and me_id:
            spouse_map.setdefault(bo_id, set()).add(me_id)
            spouse_map.setdefault(me_id, set()).add(bo_id)

def get_parent_uuid(parent_name, child_life):
    if pd.isna(parent_name) or str(parent_name).strip() == '': return ''
    p_name = str(parent_name).strip()
    try: p_life = safe_int_convert(child_life) - 1 if safe_int_convert(child_life) else ''
    except: p_life = ''
    return name_to_uuid.get(f"{p_name}|{p_life}", '')

output_columns = [
    'No', 'Mã thành viên', 'Đời', 'Chi', 'Nhánh', 'Ngành', 'Phái', 'Cành', 'Hệ', 'Mã dòng họ', 'STT',
    'Tên', 'Tên khác', 'LK bổ xung', 'Giới tính', 'Nội tộc', 'Thứ bậc', 'Tóm lược', 
    'Dữ liệu sinh tự động', 'Dữ liệu cần cập nhật', 'Đầy đủ', 'Ngày sinh dương lịch', 'Ngày sinh âm lịch', 
    'Giờ sinh can chi', 'Ngày sinh can chi', 'Tháng sinh can chi', 'Năm sinh can chi',
    'Đã mất', 'Ngày giỗ dương lịch', 'Ngày giỗ âm lịch', 'Giờ mất can chi', 'Ngày mất can chi', 
    'Tháng mất can chi', 'Năm mất can chi', 'Tuổi thọ', 'Mộ phần', 'Toạ độ mộ phần', 'Nghĩa trang', 
    'Địa chỉ nghĩa trang', 'Nơi thờ cúng', 'Người thờ cúng', 'Mã thành viên bố', 'Tên bố', 'Mã thành viên mẹ', 
    'Tên mẹ', 'Mã thành viên chồng', 'Tên chồng', 'Mã thành viên vợ', 'Tên vợ', 'Tỉnh/ thành phố quê quán', 
    'Quận/ huyện quê quán', 'Phường/ xã quê quán', 'Số nhà / đường quê quán', 'Quê quán', 'Tỉnh/ thành phố thường trú', 
    'Quận/ huyện thường trú', 'Phường/ xã thường trú', 'Số nhà / đường thường trú', 'Thường trú',
    'Định vị', 'Địa chỉ hoạt động', 'Ngành nghề', 'Chi tiết ngành nghề', 'Chức vụ', 'Đơn vị công tác',
    'Mô tả khách hàng mục tiêu', 'Mô tả danh mục sản phẩm dịch vụ', 'Trình độ học vấn', 'Trình độ chuyên môn', 
    'Trường đào tạo', 'Khoa', 'Khóa', 'Lớp', 'Chuyên ngành', 'Tình trạng hôn nhân', 'Đảng viên', 
    'Có công với đất nước', 'Dân tộc', 'Quốc tịch', 'Username', 'Email', 'Password', 'Điện thoại', 
    'Zalo', 'Facebook', 'Youtube', 'Tiktok', 'Website', 'Thư viện', 'Ẩn', 'Hiển thị tt cá nhân', 'Liệt sỹ', 
    'Thương binh', 'Bệnh binh', 'Mẹ Việt Nam anh hùng', 'Huân huy chương', 'Lực lượng vũ trang nhân dân', 
    'Đỗ khoa bảng', 'Thành đạt trên lĩnh vực', 'Quan chức tước', 'Cống hiến cho dòng họ',
    'Đường dẫn ảnh chân dung', 'Đường dẫn ảnh mộ phần', 'Đường dẫn ảnh thành tích', 'Đường dẫn ảnh người có công'
]

output = pd.DataFrame(columns=output_columns)
output['No'] = df['STT']
output['STT'] = df['STT']
output['Mã thành viên'] = df['Mã thành viên']
output['Tên'] = df['Tên hiển thị'] 
output['Giới tính'] = df['Giới tính']
output['Đời'] = df['Đời']
output['Mã dòng họ'] = df['Mã dòng họ']
output['Thứ bậc'] = df['Thứ bậc']
output['Nội tộc'] = df['Nội tộc']
output['Dữ liệu sinh tự động'] = 'X'

for i, row in df.iterrows():
    m_id = row['Mã thành viên']
    gt = str(row['Giới tính']).strip().upper()
    spouses = list(spouse_map.get(m_id, []))
    sp_names = [str(uuid_to_raw_name.get(s_id, "")).strip() for s_id in spouses]
    
    if gt == 'NAM':
        output.at[i, 'Mã thành viên vợ'] = ", ".join(spouses)
        output.at[i, 'Tên vợ'] = ", ".join(sp_names)
    elif gt == 'NỮ':
        output.at[i, 'Mã thành viên chồng'] = ", ".join(spouses)
        output.at[i, 'Tên chồng'] = ", ".join(sp_names)

output['Đầy đủ'] = df.apply(lambda r: build_full_description(r, df), axis=1)

output['Ngày sinh dương lịch'] = df.get('Ngày sinh (DL)', '').apply(standardize_date)
output['Ngày sinh âm lịch']    = df.get('Ngày sinh (AL)', '').fillna('').astype(str).str.strip()
output['Ngày giỗ dương lịch']  = df.get('Ngày mất (DL)', '').apply(standardize_date)
output['Ngày giỗ âm lịch']     = df.get('Ngày mất (AL)', '').fillna('').astype(str).str.strip()
output['Đã mất'] = df['Trạng thái'].apply(lambda x: 'X' if str(x).strip() == 'Đã mất' else '')

output['Tên bố'] = df.get('Tên bố', '').fillna('').astype(str).str.strip()
output['Mã thành viên bố'] = df.apply(lambda r: get_parent_uuid(r.get('Tên bố'), r.get('Đời')), axis=1)
output['Tên mẹ'] = df.get('Tên mẹ', '').fillna('').astype(str).str.strip()
output['Mã thành viên mẹ'] = df.apply(lambda r: get_parent_uuid(r.get('Tên mẹ'), r.get('Đời')), axis=1)

output['Quốc tịch'] = df.get('Quốc tịch', 'Việt Nam').fillna('Việt Nam')
output['Dân tộc'] = 'Kinh'
output['Quê quán'] = df.get('Nguyên quán', '').fillna('')
output['Tỉnh/ thành phố quê quán'] = df.get('Nguyên quán', '').fillna('')

for col in output.columns:
    output[col] = output[col].apply(lambda x: '' if pd.isna(x) or str(x).lower() == 'nan' else x)

output_file = 'Output_Final.xlsx'
output.to_excel(output_file, index=False)