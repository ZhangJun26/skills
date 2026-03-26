# SlideBlocks — PPT 智能组装 Skill

告诉 AI 你要做什么，它帮你组装 PPT、换模板风格、或整理已有幻灯片。

---

## 两种使用模式

### Mode A：PPT 工具（无需素材库）

只需要本地有 PPT 文件，直接跟 AI 说：

| 功能 | 说明 | 示例 |
|------|------|------|
| **统一标题栏** | 给已有 PPT 批量套用新模板的标题栏风格 | "把这个 PPT 换成科技风深色模板" |
| **深浅色转换** | 整份 PPT 从深色底换浅色底（或反向） | "帮我把这份 PPT 转成浅色底版本" |
| **局部编辑** | 删页、移页、插过渡页、替换某页 | "删掉第 3 页，在第 5 章前加一个过渡页" |

**Mode A 不需要配置 `materials_dir`，也不需要运行 `setup_paths.py`。** 只需填好 `output_dir` 即可。

---

### Mode B：素材库搜索 + 组装（需要素材库）

从素材库里搜页面、选内容、套模板、组装成新 PPT。

> "帮我做一个售前汇报，客户是医院，主题是 AI 辅助诊断，需要背景、产品、案例三个章节"

素材库有两种来源（二选一）：

**① 直接用现有库**（推荐）

找维护人要 `slide_vault.db` 和素材文件夹（或网络共享盘路径）。
按下方「配置步骤」完成后即可使用。

**② 自建素材库**

指向自己本地的 PPT 文件夹，扫描打标签后就有专属素材库：

```bash
# 1. 扫描 PPT，建立索引
python slide_vault/scanner.py

# 2. 自动打标签（会调用 Claude API，需要配置 ANTHROPIC_API_KEY）
python tools/auto_tag.py

# 3. 完成，可以开始搜索组装
```

---

## 环境要求

- Windows 系统
- Microsoft PowerPoint 或 WPS（必须安装）
- Python 3.9+

```bash
pip install anthropic python-pptx pyyaml flask pywin32
```

---

## 配置步骤（Mode B 必须；Mode A 只需填 output_dir）

### 第一步：填写 config.yaml

```yaml
materials_dir: "素材文件夹的本地路径"  # Mode A 可留空；Mode B 必填（填下载后的素材文件夹）
output_dir:    "你本地的输出文件夹"    # 两种模式都必填，建议放在 skill 目录之外
```

> `output_dir` 建议填 skill 目录之外的路径（如 `D:/PPT输出`），避免 skill 更新时被覆盖。

### 第二步：初始化数据库路径（Mode B 必须，只跑一次）

```bash
python setup_paths.py
```

---

## 模板

`模板/` 文件夹里有两套默认模板：
- **科技风（深色底）** — 深色背景，适合 AI / 数字化方向
- **蓝色商务（浅色底）** — 白色背景，通用商务风格

想加自己的模板，见 `模板/模板说明.md`。

---

## 素材库有更新怎么办

让维护人把新的 `slide_vault.db` 发给你，替换掉 skill 文件夹里的同名文件，再重新运行一次 `setup_paths.py`。

---

## 常见问题

**组装时 PowerPoint 窗口一直弹出来？**
正常现象，完成后自动关闭。期间不要手动点 PowerPoint 窗口。

**搜索结果不理想？**
换几个同义词试试，或告诉 AI "换个关键词再搜一次"。

**电脑同时装了 WPS 和 Office？**
会优先调用 Office，只有 WPS 也能正常用。
