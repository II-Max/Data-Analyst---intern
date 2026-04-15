"""
Script c?u h?nh d? li?u cho ?ng d?ng
"""

# C?u h?nh c?t d? li?u - tùy ch?nh theo nhu c?u c?a b?n
COLUMNS_CONFIG = [
    {
        'name': 'STT',
        'type': 'number',
        'required': True,
        'width': 60,
        'format': '0',
        'description': 'S? th? t?'
    },
    {
        'name': 'Tên',
        'type': 'text',
        'required': True,
        'width': 150,
        'description': 'H? và tên'
    },
    {
        'name': 'Gi?i tính',
        'type': 'select',
        'required': True,
        'options': ['Nam', 'N?'],
        'width': 100,
        'description': 'Gi?i tính'
    },
    {
        'name': 'Nãm sinh',
        'type': 'number',
        'required': False,
        'width': 100,
        'format': '0',
        'description': 'Nãm sinh (YYYY)'
    },
    {
        'name': 'Ngày sinh',
        'type': 'date',
        'required': False,
        'width': 120,
        'format': 'dd/mm/yyyy',
        'description': 'Ngày sinh (DD/MM/YYYY)'
    },
    {
        'name': 'Nõi sinh',
        'type': 'text',
        'required': False,
        'width': 150,
        'description': 'Nõi sinh'
    },
    {
        'name': 'Tên b?',
        'type': 'text',
        'required': False,
        'width': 150,
        'description': 'Tên cha'
    },
    {
        'name': 'Tên m?',
        'type': 'text',
        'required': False,
        'width': 150,
        'description': 'Tên m?'
    },
    {
        'name': 'Ð?a ch?',
        'type': 'text',
        'required': False,
        'width': 250,
        'description': 'Ð?a ch? hi?n t?i'
    },
    {
        'name': 'S? ði?n tho?i',
        'type': 'text',
        'required': False,
        'width': 120,
        'description': 'S? ði?n tho?i'
    },
    {
        'name': 'Email',
        'type': 'email',
        'required': False,
        'width': 180,
        'description': 'Email'
    },
    {
        'name': 'Ghi chú',
        'type': 'text',
        'required': False,
        'width': 200,
        'description': 'Ghi chú thêm'
    },
]

# C?u h?nh Excel
EXCEL_CONFIG = {
    'header_color': '4472C4',      # Màu header (Blue)
    'header_text_color': 'FFFFFF', # Màu text header (White)
    'header_font_size': 11,
    'header_bold': True,
    'row_height': 20,
    'header_height': 25,
    'border_color': '000000',
    'border_style': 'thin',
    'default_font': 'Calibri',
    'sheet_name': 'Sheet1',
}

# C?u h?nh ?ng d?ng
APP_CONFIG = {
    'debug': True,
    'host': '0.0.0.0',
    'port': 5000,
    'max_content_length': 16 * 1024 * 1024,  # 16MB max file size
    'upload_folder': 'uploads',
    'allowed_extensions': {'xlsx'},
}

# Validation rules
VALIDATION_RULES = {
    'STT': {
        'type': 'number',
        'min': 1,
        'max': 10000,
    },
    'Nãm sinh': {
        'type': 'number',
        'min': 1900,
        'max': 2024,
    },
    'S? ði?n tho?i': {
        'type': 'phone',
        'pattern': r'^\d{10,11}$',
    },
    'Email': {
        'type': 'email',
    }
}
