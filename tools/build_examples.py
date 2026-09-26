"""Rebuild public examples from bundled sample vocabulary, never user uploads."""
from pathlib import Path
import sys
import tempfile
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'skills/ebbinghaus-vocabulary/scripts'))
import vocab as v


def main():
    words, _ = v.load_words(ROOT / 'skills/ebbinghaus-vocabulary/examples/词表示例.csv')
    examples = ROOT / 'examples'
    cases = [('01-中文英文', ()), ('02-中文英文词性', ('pos',)),
             ('03-中文英文音标', ('ipa',)), ('04-完整四字段', ('ipa', 'pos'))]
    counts = {}
    with tempfile.TemporaryDirectory() as temp:
        temp = Path(temp)
        for name, optional in cases:
            data = [{k: value for k, value in w.items() if k in ('english', 'chinese') + optional} for w in words]
            for daily in (5, 10):
                work = temp / name / f'每组{daily}词'
                work.mkdir(parents=True)
                plan = v.schedule(data, daily)
                v.render_xlsx(v.master_layout(plan['groups']), work / '单词总表.xlsx', work)
                if not optional:
                    target = examples / f'300-words-{daily}-per-day'
                    target.mkdir(parents=True, exist_ok=True)
                    (target / '单词总表.xlsx').write_bytes((work / '单词总表.xlsx').read_bytes())
                    pages = v.pdf_layout(plan)
                    v.render_pdf(pages, target / '艾宾浩斯复习打印册.pdf')
                    counts[daily] = len(pages)
        mixed = [{k: val for k, val in w.items() if k in ('english', 'chinese') + cases[i // 10][1]}
                 for i, w in enumerate(words[:40])]
        mixed[31].pop('ipa', None)
        mixed[32].pop('pos', None)
        work = temp / '05-同页混合字段与部分缺失' / '每组10词'
        work.mkdir(parents=True)
        v.render_xlsx(v.master_layout(v.schedule(mixed, 10)['groups']), work / '单词总表.xlsx', work)
        with zipfile.ZipFile(examples / 'adaptive-layout-samples.zip', 'w', zipfile.ZIP_DEFLATED) as archive:
            for path in sorted(temp.rglob('单词总表.xlsx')):
                archive.write(path, path.relative_to(temp).as_posix())
            archive.writestr('测试说明.txt', 'v1.0.1：9份样例覆盖2/3/4字段、5/10词、混合字段；字段标题与内容左对齐，A4纵向，完整组不拆分。')
    readme = examples / 'README.md'
    import re
    text = readme.read_text(encoding='utf-8')
    text = re.sub(r'(5词方案：.*?PDF)\d+(页)', rf'\g<1>{counts[5]}\2', text)
    text = re.sub(r'(10词方案：.*?PDF)\d+(页)', rf'\g<1>{counts[10]}\2', text)
    if '以下文件已使用v1.0.1重新生成' not in text:
        text = text.replace('样例使用技能包内', '以下文件已使用v1.0.1重新生成，字段标题统一左对齐。样例使用技能包内')
    readme.write_text(text, encoding='utf-8')


if __name__ == '__main__':
    main()
