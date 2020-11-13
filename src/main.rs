use anyhow::{Context, Result};
use std::cell::RefCell;
use std::collections::HashMap;
use std::fs;
use std::path::Path;
use std::rc::{Rc, Weak};
use structopt::StructOpt;
use xml::{
    attribute::OwnedAttribute,
    name::OwnedName,
    reader::{EventReader, XmlEvent},
};

#[derive(Debug, StructOpt)]
#[structopt(name = "ooxml", about = "An example of parsing docx")]
struct Opt {
    /// Specify file name of .docx, I.E. demo.docx
    #[structopt()]
    file_name: String,

    /// Activate verbose mode
    #[structopt(short, long)]
    verbose: bool,
}

/// 简单版本:
/// 运行
/// ```sh
/// cargo run -- demo.docx
/// ```
/// 输出字体,并且带字体的颜色值.
fn main() -> Result<()> {
    let opt = Opt::from_args();

    let file_name = Path::new(&opt.file_name);
    let file =
        fs::File::open(file_name).with_context(|| format!("open file {:?} err", file_name))?;

    // 使用 zip 创建该文件的 Archive
    let mut archive = zip::ZipArchive::new(file).context("create zip archive err")?;
    for i in 0..archive.len() {
        let file = archive.by_index(i).unwrap();
        if opt.verbose {
            println!("filename: {}", file.name());
        }
    }

    // 直接解析 main document: word/document.xml
    // TODO 这个是写死的路径,正常应该先解析 [Content_types].xml 找到 main document.
    let word_doc = archive
        .by_name("word/document.xml")
        .context("found no word/document.xml")?;

    // xml parse
    let parser = EventReader::new(word_doc);
    let mut depth = 0;
    for e in parser {
        let event = e.context("xml parser got err")?;
        match event {
            XmlEvent::StartElement {
                name,
                attributes,
                namespace: _,
            } => {
                if opt.verbose {
                    print_xml_owned_name(&name, depth, true);
                }
                if name.local_name.ne("r") {
                    let attrs: Vec<_> = attributes
                        .iter()
                        .map(|a| format!("{}={}", a.name.local_name, a.value))
                        .collect();
                    println!(
                        "{}{}, attributes: {:?}",
                        " ".repeat(depth),
                        name.local_name,
                        attrs
                    );
                }
                depth += 1;
            }
            XmlEvent::EndElement { name } => {
                depth -= 1;
                if opt.verbose {
                    print_xml_owned_name(&name, depth, false);
                }
            }
            XmlEvent::Comment(ref data) => {
                println!(r#"Comment("{}")"#, data.escape_debug());
            }
            XmlEvent::CData(ref data) => {
                println!(r#"CData("{}")"#, data.escape_debug());
            }
            XmlEvent::Characters(ref data) => {
                println!(
                    r#"{}Characters("{}")"#,
                    " ".repeat(depth),
                    data.escape_debug()
                );
            }
            XmlEvent::Whitespace(ref data) => {
                println!(r#"Whitespace("{}")"#, data.escape_debug());
            }
            _ => {
                // TODO
            }
        }
    }
    Ok(())
}

fn print_xml_owned_name(name: &OwnedName, indent: usize, start: bool) {
    print!("{}", " ".repeat(indent));
    if start {
        print!("+");
    } else {
        print!("-");
    }
    if let Some(v) = &name.prefix {
        print!("{}:", v);
    }
    println!("{}", name.local_name);
}

struct Paragraph {
    property: Option<ParagraphProperty>,
    runs: Vec<Run>,
}

struct ParagraphProperty {
    // paragraph 的默认 run property
    run_property: Option<RunContentColor>,
    // TODO 省略其他..
}

struct Run {
    property: RunProperty,
    contents: Vec<RunContent>,
}

struct RunProperty {
    bold: bool,
    italic: bool,
    color: Option<RunContentColor>,
    // TODO 省略其他..
}

/// Run content color property.
struct RunContentColor {
    val: Option<String>,
    theme_color: Option<String>,
    // TODO 省略其他..
}

enum ElementType {
    Document,
    Body,
    Paragraph,
    Run,
    Text,
    /// 属性
    Property,

    /// 其他剩余的不支持的类型
    Unknown,
}
impl ElementType {
    fn from_str(s: &str) -> Self {
        match s {
            "document" => Self::Document,
            "body" => Self::Body,
            "p" => Self::Paragraph,
            "r" => Self::Run,
            "t" => Self::Text,
            "pPr" => Self::Property,
            "rPr" => Self::Property,
            _ => Self::Unknown,
        }
    }
}
struct Element {
    element_type: ElementType,
    parent: Option<Weak<RefCell<Element>>>,
    children: Vec<Rc<RefCell<Element>>>,
    attributes: HashMap<String, String>,
    literal_text: Option<String>, // 目前只有  w:t 有
}
impl Element {
    fn new(element_type: ElementType) -> Self {
        Self {
            element_type,
            parent: None,
            children: vec![],
            attributes: HashMap::new(),
            literal_text: None,
        }
    }
}

struct Parsing {
    // 这里假设有一个唯一的 root
    root: Option<Rc<RefCell<Element>>>,
    current: Option<Rc<RefCell<Element>>>,
}
impl Parsing {
    fn feed_element(&mut self, name: OwnedName, attributes: Vec<OwnedAttribute>) {
        if name.prefix.ne("w") {
            println!("unknown prefix of name: {:?}", name);
            return;
        }

        let element_type = ElementType::from_str(&name.local_name);
        let element = Rc::new(RefCell::new(Element::new(element_type)));
        if self.root.is_none() {
            // 第一个节点
            self.root.replace(Rc::clone(&element));
            self.current.replace(element);
            return;
        }
        // 平级添加? 还是添加到 children 里
    }
    fn fish_feed_element() {}
    /// 目前只有 w:t 类型会有
    fn feed_characters(&mut self, data: String) {}
}
