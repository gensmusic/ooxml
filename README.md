# ooxml

## 目录结构

```shell script
ooxml.md    # OOXML的介绍文档
src/main.rs # 解析 docx 文档并且打印字体的颜色和内容.
result.jpg  # 上面程序运行的一个示例截图.

OOXML.xmind # ooxml 资料的脑图(xmind 打开)
OOXML.png   # ooxml 脑图的图片格式

misc/*.pdf  # ECMA-376 5th 的 part1, part4资料
demo.docx   # 一个 docx 的测试文档
demo        # demo.docx 使用 zip 解压缩后的目录
```

## 程序运行

在程序根目录下运行,接收一个参数 `docx`文件名.

```shell script
cargo run -- demo.docx

# 查看其它参数
cargo run -- -h
```