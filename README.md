# Tuomin - Excel 脱敏工具

命令行工具，对 xls/xlsx 文件中的个人敏感信息进行脱敏处理。

## 安装

```bash
pip install tuomin
```

Python >= 3.8

### 开发

```bash
git clone https://github.com/annopick/tuomin.git
cd tuomin
uv sync
```

## 用法

```bash
tuomin <input_file> --col <列字母>:<字段类型>:<脱敏方式> [--col ...] [选项]
```

### 参数

| 参数 | 说明 |
|---|---|
| `input_file` | 输入的 xls/xlsx 文件路径 |
| `--sheet` | Sheet 名称或索引（从 0 开始），默认第一个 Sheet |
| `--col` | 指定脱敏列，格式 `列字母:字段类型:脱敏方式`，可多次使用 |
| `--output` | 输出文件路径，默认在同目录生成 `_masked.xlsx` 后缀文件 |

### 字段类型

| 标识 | 说明 |
|---|---|
| `name` | 姓名 |
| `phone` | 手机号 |
| `idcard` | 身份证号 |
| `email` | 邮箱 |
| `bankcard` | 银行卡号 |
| `address` | 地址 |

### 脱敏方式

| 标识 | 说明 |
|---|---|
| `replace` | 用 `*` 替换敏感部分，保留部分原文特征 |
| `reconstruct` | 生成与现实无关的假数据 |

## 示例

替换单列手机号：

```bash
tuomin data.xlsx --col C:phone:replace
```

重构多列：

```bash
tuomin data.xls --col B:name:reconstruct --col D:idcard:reconstruct --output result.xlsx
```

指定 Sheet + 混合脱敏：

```bash
tuomin data.xlsx --sheet 员工信息 --col B:name:replace --col C:phone:reconstruct --col E:email:replace
```

## 替换规则

| 字段类型 | 保留规则 | 示例 |
|---|---|---|
| 姓名 | 保留姓，名用 `*` | 张三丰 → `张**` |
| 手机号 | 保留前 3 后 4 | 13812341234 → `138****1234` |
| 身份证号 | 保留前 3 后 4 | 310101199001011234 → `310***********1234` |
| 邮箱 | 首字符 + `*` + 域名 | zhangsan@qq.com → `z*******@qq.com` |
| 银行卡号 | 保留前 4 后 4 | 6222021234565678 → `6222********5678` |
| 地址 | 保留到区/县 | 上海市浦东新区张江路88号 → `上海市浦东新区******` |

## 重构规则

生成的假数据不可能对应任何真实信息：

| 字段类型 | 生成规则 | 示例 |
|---|---|---|
| 姓名 | 姓氏 + 中文数字 | 张三四 |
| 手机号 | 110 开头（报警号段） | 11056781234 |
| 身份证号 | 999 开头（不存在的地区码） | 999999199901019923 |
| 邮箱 | @notexist.invalid | a3k9x@notexist.invalid |
| 银行卡号 | 0000 开头（不存在的 BIN） | 0000123456781234 |
| 地址 | 虚构路 + 数字号 | 广东省深圳市虚构路999号 |

## 特性

- **一致性** — 同一原文在同一列中始终映射到相同结果
- **叠加重构** — 对已包含 `*` 的替换数据可再次重构
- **空值/非标准值** — 原样保留，不报错
- **编码兼容** — 自动处理 xls GBK 编码和控制台输出编码
- **安全** — 永不覆盖原文件，始终生成新文件
- **格式统一** — 输出始终为 xlsx
