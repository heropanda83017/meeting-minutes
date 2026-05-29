# python-docx 数字字体设置避坑指南

## 问题背景

在会议记录排版中，所有阿拉伯数字需使用 Times New Roman 字体，与中文仿宋/黑体/楷体区分。这在 Word 中是一个常见的排版需求，但在 python-docx 中实现时存在陷阱。

## 坑：p.add_run() 的迭代器冲突

### ❌ 错误做法

```python
for run in p.runs:          # 遍历 runs
    parts = text.split(...)
    run._element.clear()     # 清空原 run
    for part in parts:
        new_run = p.add_run(part)  # 添加新 run 到段落
```

**问题**：`p.add_run()` 会在运行中修改 `p.runs` 列表，导致迭代器计数混乱。
- 当外循环 `for run in p.runs` 正在迭代时，内层 `p.add_run()` 添加新元素到列表末尾
- 这会导致部分 run 被重复处理或跳过
- 最终结果：文字内容重复出现或错位

### ✅ 正确做法：lxml XML 级操作

使用 lxml 直接在 OOXML 级别的 DOM 树上操作 `<w:r>` 元素：

```python
from lxml import etree
from docx.oxml.ns import qn

for run in list(p.runs):            # list() 创建拷贝，避免迭代器冲突
    text = run.text
    if not re.search(r'\d', text):
        continue
    parts = re.split(r'(\d+)', text)  # 拆分为 ["中文字符", "123", "更多中文", "456"]

    # 1. 获取原 run 的字体信息
    orig_font = run.font.name
    orig_size = run.font.size
    orig_bold = run.font.bold
    rpr = run._element.find(qn('w:rPr'))
    east_asia = rpr.find(qn('w:rFonts')).get(qn('w:eastAsia')) if rpr is not None else None

    # 2. 在原位置删除 run
    parent = run._element.getparent()
    idx = list(parent).index(run._element)
    parent.remove(run._element)

    # 3. 在原地插入新 run 元素
    new_elems = []
    for part in parts:
        elem = etree.SubElement(parent, qn('w:r'))
        rp = etree.SubElement(elem, qn('w:rPr'))
        rf = etree.SubElement(rp, qn('w:rFonts'))
        if re.fullmatch(r'\d+', part):
            rf.set(qn('w:ascii'), 'Times New Roman')
            rf.set(qn('w:hAnsi'), 'Times New Roman')
            rf.set(qn('w:eastAsia'), 'Times New Roman')
        else:
            rf.set(qn('w:ascii'), orig_font or '仿宋')
            rf.set(qn('w:eastAsia'), east_asia or orig_font or '仿宋')
        # 保留原字号和加粗
        if orig_size:
            sv = str(int(orig_size.pt * 2))
            etree.SubElement(rp, qn('w:sz')).set(qn('w:val'), sv)
            etree.SubElement(rp, qn('w:szCs')).set(qn('w:val'), sv)
        if orig_bold:
            etree.SubElement(rp, qn('w:b'))
        t = etree.SubElement(elem, qn('w:t'))
        t.text = part
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        new_elems.append(elem)

    # 4. 调整位置：etree.SubElement 插入到末尾，需移到原位
    for ne in new_elems:
        parent.remove(ne)
    for j, ne in enumerate(new_elems):
        parent.insert(idx + j, ne)
```

## 关键原理

| 步骤 | 说明 |
|------|------|
| `list(p.runs)` | 创建 runs 的副本，避免迭代器修改问题 |
| `run._element.getparent()` | 获取段落对应的 XML `<w:p>` 元素 |
| `list(parent).index(run._element)` | 记录原 run 在段落中的位置 |
| `etree.SubElement(parent, qn('w:r'))` | 直接在 XML 级插入新 run |
| `rFonts_elem.set(qn('w:ascii'), 'Times New Roman')` | 设置数字部分的西文字体 |
| `rFonts_elem.set(qn('w:eastAsia'), ...)` | 设置中文部分的 EastAsia 字体 |
| `parent.remove(ne); parent.insert(idx+j, ne)` | 重排元素到正确顺序 |

## 字体属性对应关系

| OOXML 属性 | python-docx 属性 | 说明 |
|-----------|-----------------|------|
| `w:ascii` | `run.font.name` | 西文字体名 |
| `w:eastAsia` | (无直接映射) | 东亚字体名，RPR.rFonts.get('w:eastAsia') |
| `w:hAnsi` | (无直接映射) | 高字节 ANSI 字体名 |
| `w:sz` | `run.font.size` | 字号，单位=磅值×2 |
| `w:b` | `run.font.bold` | 加粗标记 |

## 注意事项

1. **`run.font.element.rPr.rFonts` 是只读属性**，不能赋值（`rPr.rFonts = ...` 会抛出 `AttributeError`）。必须使用 `.set()` 方法。
2. **字号单位**：OOXML 中的字号值是磅值×2（如 16pt → 32），python-docx 中用 `Pt(x)` 对象。
3. **保留命名空间**：所有元素操作必须使用 `qn('w:xxx')` 来添加 `w:` 命名空间前缀。
4. **混排处理**：当一段文字中数字和非数字交替出现时（如"第1阶段共296条"），一个 run 需要拆分为多个 run。
