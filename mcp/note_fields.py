"""Human-editable task fields stored as ordinary Project Notes text."""
import re
import math

START = '【MCP任务字段:v1】'
END = '【/MCP任务字段】'
BLOCK = re.compile(re.escape(START) + r'\r?\n(.*?)\r?\n' + re.escape(END), re.S)
LINE = re.compile(r'^\s*([^：:\r\n]{1,40})\s*[：:]\s*(.*?)\s*$')


def read_fields(notes: str) -> dict:
    normalized = (notes or '').replace('\r\n', '\n').replace('\r', '\n')
    blocks = list(BLOCK.finditer(normalized))
    if (START in (notes or '') or END in (notes or '')) and not blocks:
        raise ValueError('任务字段区标记不完整，请补全后再修改')
    if len(blocks) > 1:
        raise ValueError('存在多个 MCP 任务字段区，请合并后再修改')
    if normalized.count(START) != len(blocks) or normalized.count(END) != len(blocks):
        raise ValueError('任务字段区标记不成对')
    text = blocks[0].group(1) if blocks else normalized
    fields, duplicates, unparsed = {}, [], []
    for line in text.splitlines():
        match = LINE.fullmatch(line)
        if not match:
            if line.strip():
                unparsed.append(line)
            continue
        key, value = match[1].strip(), match[2]
        if key in fields:
            duplicates.append(key)
        fields[key] = value
    return {'fields': fields, 'duplicate_keys': duplicates, 'managed_block': bool(blocks),
            'notes': notes or '', 'unparsed_lines': unparsed}


def merge_fields(notes: str, updates: dict) -> str:
    notes = notes.replace('\r\n', '\n').replace('\r', '\n')
    parsed = read_fields(notes)
    if parsed['duplicate_keys']:
        raise ValueError(f"备注存在重名字段: {parsed['duplicate_keys']}，请先消除歧义")
    if parsed['managed_block'] and parsed['unparsed_lines']:
        raise ValueError('字段区内存在非字段文字，请先移到字段区外，防止更新时丢失')
    # For unmarked notes, preserve the whole original.  The managed block
    # starts with its existing key/value lines to support gradual migration.
    fields = dict(parsed['fields'])
    for key, value in updates.items():
        if not isinstance(key, str) or not key.strip() or len(key) > 40 or re.search(r'[\r\n：:]', key):
            raise ValueError('字段名须为 1~40 字符，不能含冒号或换行')
        if isinstance(value, (dict, list)) or value is None:
            raise ValueError('字段值须为单行文字或数字；暂不支持删除字段')
        value = str(value)
        if '\n' in value or '\r' in value or START in value or END in value:
            raise ValueError('字段值不能包含换行或字段区标记')
        if key.strip() in ('工程量', '已完成量'):
            try:
                if not math.isfinite(float(value)) or float(value) < 0:
                    raise ValueError('数量不能为负数')
            except ValueError:
                raise ValueError('工程量/已完成量请填写非负数字，单位单独填写')
        fields[key.strip()] = value
    block = START + '\r\n' + '\r\n'.join(f'{key}：{value}' for key, value in fields.items()) + '\r\n' + END
    if parsed['managed_block']:
        return BLOCK.sub(lambda _: block, notes, count=1)
    prose = '\r\n'.join(line for line in notes.splitlines() if not LINE.fullmatch(line)).rstrip()
    return (prose + '\r\n\r\n' if prose.strip() else '') + block
