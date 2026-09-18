# 排版参数与页数调优

## 模板规格

`assets/template.docx` 已调好的参数（A4）：

| 项 | 值（DXA / twips） | 说明 |
|---|---|---|
| 页面 | 11906 × 16838 | A4 纵向 |
| 页边距 | top 430 / right 782 / bottom 420 / left 782 | 约 0.30" / 0.54" |
| 内容宽 | 10342 | 页面宽 - 左右边距 |

1440 twips = 1 英寸。1 pt = 20 twips。

### 样式 ID 对照

模板使用自定义样式 ID，改排版时直接改 `word/styles.xml`：

| ID | 名称 | 用途 | 原始 spacing | 收紧后 |
|---|---|---|---|---|
| 164 | Resume Section | 章节标题（教育背景、实习经历） | before 90 / after 44 | before 60 / after 32 |
| 165 | Resume Entry | 学校行、公司行、项目标题行 | before 44 / after 10 / line 240 | before 32 / after 8 |
| 166 | Resume Meta | 角色行、技术栈行（青色小字） | after 22 / line 240 | after 16 |
| 167 | Resume Bullet | bullet 列表项 | after 35 / line 257 | after 22 / line 222 |
| 168 | Resume Compact | 普通正文段（竞赛、奖学金、项目简介） | after 27 / line 252 | after 18 / line 224 |

行距 `line` 是相对值，240 = 单倍行距。

### 字号与配色

| 元素 | sz 值 | 实际字号 | 颜色 |
|---|---|---|---|
| 姓名 | 44 | 22pt | `12324A` 深蓝 |
| 方向标签 | 23 | 11.5pt | `0C7C86` 青绿 |
| 章节标题 | 23 | 11.5pt | `12324A` |
| 条目标题行 | 20 | 10pt | `12324A` |
| 正文 / bullet | 19 | 9.5pt | `17212B` |
| 日期、次要信息 | 18 | 9pt | `5D6872` 灰 |
| 角色行（青绿） | 17 | 8.5pt | `0C7C86` |

- 主色 `0C7C86`，深色 `12324A`，正文 `17212B`，强调底色 `EFF7F7`
- 中文字体：微软雅黑（`w:eastAsia="微软雅黑"`），西文 Arial

---

## 页数调优

**目标：1 页**（应届生标准）。

### 调优顺序（重要）

按代价从低到高，**能靠前的就不动后面的**：

1. **删无关内容** —— 代价最低，收益最高
2. **合并罗列型条目** —— 例如把三条同类的技能合并成一条
3. **压缩措辞** —— 去掉"进行""相关""等"这类废字
4. **调段间距**（styles.xml 的 after 值）—— 视觉影响小
5. **调页边距**（pgMar）—— 太小会显得挤
6. **调行距**（styles.xml 的 line 值）—— 影响可读性
7. **调字号** —— 最后手段，会明显改变观感

**不要一上来就调行距和字号**。先砍内容，很多时候砍完就够了。

### 行距的硬约束

行高 ≈ 字号 × 1.2 × (line / 240)。

对 9.5pt 的正文字（sz 19）：

| line 值 | 行高 | 状态 |
|---|---|---|
| 257 | 12.2 pt | 宽松 |
| 240 | 11.4 pt | 正常 |
| 222 | 10.5 pt | 紧凑但可读 |
| 208 | 9.9 pt | 偏挤 |
| 200 | 9.5 pt | **临界**，等于字号 |
| < 195 | < 9.3 pt | **文字会重叠，不要用** |

**行高不能小于字号。** 到了 200 还想省空间，只能删内容。

### 迭代方法

改排版后要重跑验证（见 SKILL.md）。迭代节奏建议：

```
改 → 打包 → 测页数 → 渲染看图 → 看底部剩余空间
```

`render_pdf.py` 会报告首页底部剩余空间：

- 剩余 < 0：溢出了，必须砍
- 剩余 5-20 pt：刚好，别再加东西
- 剩余 > 40 pt：还能塞 3-4 行，或者可以放松行距让排版更透气

---

## 常见故障

### pack 时报 Dublin Core schema 错误

```
Failed to parse the XML resource 'http://dublincore.org/schemas/xmls/qdc/2003/04/02/dc.xsd'
```

校验器联网失败，**和你的改动无关**。用 `--validate false` 跳过，但必须自己验证：

```bash
python -c "
import re
from xml.etree import ElementTree as ET
s = open('word/document.xml', encoding='utf-8').read()
ET.fromstring(s)   # XML 语法
print('OK')
ids = re.findall(r'w14:paraId=\"([0-9A-Fa-f]+)\"', s)
print('paraId 唯一:', len(ids) == len(set(ids)))
print('w:p 配对:', len(re.findall(r'<w:p[ >]', s)) == s.count('</w:p>'))
"
```

### Word COM 报 ComputeStatistics 失败

说明命中了残留的、已卡死的 Word 进程。解决方案：

1. 用 `DispatchEx` 而不是 `Dispatch`（强制新实例，`scripts/page_count.py` 已经这么做）
2. 仍然失败就换 WPS：`KWps.Application`
3. 实在不行重启 Word 或等几分钟

**不要贸然 `taskkill WINWORD.EXE`** —— 用户可能开着未保存的文档。

### 保存时 Permission denied / Device or resource busy

目标文件被占用（用户在 Word/WPS 里开着）。两个选择：

1. 关掉再存
2. 换一个文件名存，然后告诉用户新旧两个文件的关系

**不要强行覆盖**。

### grep 关键词匹配不上，但内容明明在

**先分清是「检查」还是「改动」时遇到的**：

- **检查时**（跑 grep / 脚本核对内容在不在）：PDF 文本提取会在任意位置断行，长关键词会被拆开。用 `check_residue.py`，它会把空白全部去掉再匹配。
- **改动时**（拿匹配串去 `replace`，报 0 次）：是 XML 的 run 拆分问题，不是断行问题 —— 见下面《改文本》里那一节。

---

## XML 编辑要点

### 定位段落

用 `w14:paraId` 作为锚点（8 位十六进制，唯一）。查看全文段落：

```bash
python -c "
import re
s = open('word/document.xml', encoding='utf-8').read()
for pid, body in re.findall(r'<w:p [^>]*w14:paraId=\"([0-9A-F]+)\">(.*?)</w:p>', s, re.S):
    style = re.search(r'<w:pStyle w:val=\"(\d+)\"', body)
    txt = ''.join(re.findall(r'<w:t[^>]*>([^<]*)</w:t>', body))
    print(pid, style.group(1) if style else '-', txt[:60])
"
```

### 删段落

匹配从 `<w:p w14:paraId="XXX">` 到 `</w:p>` 的整块删除。批量删除时以**下一个段落的开头**作为锚点结尾，避免误删：

```python
pattern = r'\s*<w:p [^>]*w14:paraId="%s">.*?</w:p>' % pid
s, n = re.subn(pattern, "", s, flags=re.S)  # n 必须是 1
```

**每次都要检查 `n == 1`** —— 匹配到 0 次说明锚点写错了，匹配到多次说明 paraId 不唯一。

### 插入段落

复制相邻段落的结构，换一个未使用的 paraId（随便取一个 8 位十六进制，如 `5A1B2C3D`）。

### 改文本

替换 `<w:t>...</w:t>` 里的内容即可，`<w:rPr>` 里的格式会自动沿用。

⚠️ 如果新内容包含引号，用 XML 实体：`&#x201C;` `&#x201D;`（中文弯引号）。

#### ⚠️ 匹配串必须从 XML 里取，不能从文本提取的结果里取

**这是最容易重复踩的坑。**

Word 会把一句话拆成多个 run（换行、编辑历史、拼写检查、保存重排版都会触发拆分）。pandoc 或文本提取工具会把它们**拼回一个字符串**，但 **XML 里它们中间隔着 `<w:rPr>` 等标签，根本不是连续的**。拿拼接后的文本去 `str.replace()`，结果是匹配 0 次——而你会以为是内容写错了。

实测踩到的：

| 提取出来看到的 | XML 里实际的 |
|---|---|
| `GitHub 381 stars` | `<w:t>GitHub 3</w:t>` + `<w:t>81</w:t>` + `<w:t> stars / 10 forks…</w:t>` |
| `知识资产：基于 150+ 篇获奖论文…` | `<w:t>知识资产：</w:t>` + `<w:t>基于 150+ 篇获奖论文…</w:t>` |

第二种尤其隐蔽：**标签和正文几乎是必然分开的**（标签带 `<w:b/>`，正文不带），所以「标签+正文」拼起来的串永远匹配不上。改文本时**只取正文那一段**。

**动手前先 dump 目标段落的原始 XML**：

```bash
python -c "
import re
s = open('word/document.xml', encoding='utf-8').read()
i = s.find('关键词')        # 用不会出现在标签属性里的中文词定位
print(s[i-120:i+400])
"
```

**三条经验**：
1. 匹配串**越短越稳**，短到能落在单个 `<w:t>` 内
2. 标签和正文分开匹配，不要拼成一个串
3. 一条改动匹配 0 次时**不要默默跳过**——先 dump XML 看清楚。批量替换时每条都要 `n == 1`，任何一条不符就停下查原因，而不是继续往下改

### 特殊字符

- XML 里 `&` 必须写成 `&amp;`
- 带前后空格的 `<w:t>` 需要 `xml:space="preserve"`
