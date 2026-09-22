"""Build both installable archives from the same skill source (stdlib only)."""
from pathlib import Path
import hashlib
import json
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'skills' / 'ebbinghaus-vocabulary'
VERSION = '1.0.0'


def main():
    target = ROOT / 'dist'
    target.mkdir(exist_ok=True)
    files = {p.relative_to(SOURCE).as_posix(): p.read_bytes()
             for p in SOURCE.rglob('*') if p.is_file()
             and '__pycache__' not in p.parts and p.suffix not in ('.pyc', '.pyo')}
    files['LICENSE'] = (ROOT / 'LICENSE').read_bytes()
    hashes = []
    for workbuddy in (False, True):
        variant = '-workbuddy' if workbuddy else ''
        dest = target / f'ebbinghaus-vocabulary{variant}-v{VERSION}.zip'
        contents = dict(files)
        if workbuddy:
            text = contents['SKILL.md'].decode('utf-8').replace('\r\n', '\n')
            header, body = text.removeprefix('---\n').split('\n---\n', 1)
            metadata = {
                'display_name': '艾宾浩斯式单词复习',
                'display_name_en': 'Vocabulary Review Printables',
                'description_zh': '上传单词材料，按每天新增5或10词生成Excel答案表和汉译英复习打印册。',
                'description_en': 'Create a printable Excel answer sheet and spaced Chinese-to-English vocabulary worksheets.',
                'version': VERSION, 'author': 'ccylliu-cloud',
            }
            extra = '\n'.join(f'{k}: {json.dumps(v, ensure_ascii=False)}' for k, v in metadata.items())
            contents['SKILL.md'] = f'---\n{header}\n{extra}\n---\n{body}'.encode('utf-8')
        with zipfile.ZipFile(dest, 'w', zipfile.ZIP_DEFLATED) as archive:
            for name, data in sorted(contents.items()):
                prefix = '' if workbuddy else 'ebbinghaus-vocabulary/'
                item = zipfile.ZipInfo(prefix + name, (2026, 9, 22, 0, 0, 0))
                item.compress_type = zipfile.ZIP_DEFLATED
                item.external_attr = (0o100755 if name.endswith('.sh') else 0o100644) << 16
                archive.writestr(item, data)
        hashes.append(f'{hashlib.sha256(dest.read_bytes()).hexdigest()}  {dest.name}')
        print(dest.name)
    (target / 'SHA256SUMS.txt').write_text('\n'.join(hashes) + '\n', encoding='utf-8')


if __name__ == '__main__':
    main()
