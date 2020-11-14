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
    let mut doc_parsing = Parsing::new();
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
                depth += 1;
                doc_parsing.feed_element(name, attributes);
            }
            XmlEvent::EndElement { name } => {
                depth -= 1;
                if opt.verbose {
                    print_xml_owned_name(&name, depth, false);
                }
                doc_parsing.fish_feed_element();
            }
            XmlEvent::Comment(ref data) => {}
            XmlEvent::CData(ref data) => {}
            XmlEvent::Characters(data) => {
                println!(
                    r#"{}Characters("{}")"#,
                    " ".repeat(depth),
                    data.escape_debug()
                );
                doc_parsing.feed_characters(data);
            }
            XmlEvent::Whitespace(ref data) => {}
            _ => {
                // TODO
            }
        }
    }
    print_elements(&doc_parsing.root);
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

#[derive(Debug)]
enum ElementType {
    Document(String),
    Body(String),
    Paragraph(String),
    Run(String),
    Text(String),
    /// 属性
    ParagraphProperty(String),
    RunProperty(String),
    Color(String),

    /// 其他剩余的不支持的类型
    Unknown(String),
}
impl ElementType {
    fn from_name(name: &OwnedName) -> Self {
        let raw = format!(
            "{}:{}",
            name.prefix.as_ref().unwrap_or(&String::new()),
            name.local_name
        );
        if name.prefix.is_none() || name.prefix.as_ref().unwrap().ne("w") {
            return Self::Unknown(raw);
        }
        // 只匹配 w:x
        match &*name.local_name {
            "document" => Self::Document(raw),
            "body" => Self::Body(raw),
            "p" => Self::Paragraph(raw),
            "r" => Self::Run(raw),
            "t" => Self::Text(raw),
            "pPr" => Self::ParagraphProperty(raw),
            "rPr" => Self::RunProperty(raw),
            "color" => Self::Color(raw),
            _ => Self::Unknown(raw),
        }
    }
    fn is_text(&self) -> bool {
        match self {
            Self::Text(_) => true,
            _ => false,
        }
    }
    fn is_run_property(&self) -> bool {
        match self {
            Self::RunProperty(_) => true,
            _ => false,
        }
    }
    fn is_color(&self) -> bool {
        match self {
            Self::Color(_) => true,
            _ => false,
        }
    }
}
struct Element {
    element_type: ElementType,
    parent: Option<Weak<RefCell<Element>>>,
    children: Vec<Rc<RefCell<Element>>>,
    attributes: HashMap<String, String>,
    literal_text: Option<String>, // 目前只有  w:t 有
    depth: usize,                 // for debug
}
impl Element {
    /// new Element, 需要指定 parent 和 type, parent 可以为 None
    fn new(
        element_type: ElementType,
        parent: &Option<Rc<RefCell<Element>>>,
        attributes: Vec<OwnedAttribute>,
        depth: usize,
    ) -> Self {
        let mut attrs = HashMap::new();
        attributes.iter().for_each(|v| {
            attrs.insert(v.name.local_name.clone(), v.value.clone());
        });
        Self {
            element_type,
            parent: parent.as_ref().map(Rc::downgrade),
            children: vec![],
            attributes: attrs,
            literal_text: None,
            depth,
        }
    }
    fn append_child(&mut self, child: Rc<RefCell<Element>>) {
        self.children.push(child);
    }

    // 下面是一些辅助方法
    /// 寻找本节点最近的 run property
    fn find_run_property(element: &Option<Rc<RefCell<Element>>>) -> Option<Rc<RefCell<Element>>> {
        if let Some(ele) = element {
            if let Some(parent) = &ele.borrow().parent {
                if let Some(parent) = parent.upgrade() {
                    // find run property from parent's children
                    for child in parent.borrow().children.iter() {
                        if child.borrow().element_type.is_run_property() {
                            return Some(Rc::clone(child));
                        }
                    }
                    // if not found, goes up
                    return Self::find_run_property(&Some(parent));
                }
            }
        }
        None
    }

    /// 如果自己是 run property, 从中获取 color 属性
    fn get_color(element: &Option<Rc<RefCell<Element>>>) -> Option<String> {
        if let Some(ele) = &element {
            // 本身不是 run property
            if ele.borrow().element_type.is_run_property() {
                println!("+++ type not run pr");
                return None;
            }
            // 从 children 中寻找 w:color
            for child in ele.borrow().children.iter() {
                let child_ref = child.borrow();
                if child_ref.element_type.is_color() {
                    return child_ref.attributes.get("val").map(|v| v.clone());
                }
            }
            println!("+++ no color child");
        }
        println!("+++ none run pr");
        None
    }

    fn display(root: &Option<Rc<RefCell<Element>>>) -> String {
        if let Some(root_rc) = root {
            let attrs: Vec<_> = root_rc
                .borrow()
                .attributes
                .iter()
                .map(|(k, v)| format!("{}={},", k, v))
                .collect();
            let indent = "  ".repeat(root_rc.borrow().depth);
            format!(
                "{}{:?}, attrs: {:?},",
                indent,
                root_rc.borrow().element_type,
                attrs
            )
        } else {
            "None<Element>".to_string()
        }
    }
}

struct Parsing {
    // 这里假设有一个唯一的 root
    root: Option<Rc<RefCell<Element>>>,
    cur: Option<Rc<RefCell<Element>>>,
    depth: usize,
}
impl Parsing {
    fn new() -> Self {
        Self {
            root: None,
            cur: None,
            depth: 0,
        }
    }
    fn feed_element(&mut self, name: OwnedName, attributes: Vec<OwnedAttribute>) {
        self.depth += 1;

        let element_type = ElementType::from_name(&name);

        // TODO remove me
        if name.local_name.ne("r") {
            let attrs: Vec<_> = attributes
                .iter()
                .map(|a| format!("{}={}", a.name.local_name, a.value))
                .collect();
            println!(
                "{}{}, Type:{:?}, attributes: {:?}",
                " ".repeat(self.depth),
                name.local_name,
                element_type,
                attrs
            );
        }

        let element = Rc::new(RefCell::new(Element::new(
            element_type,
            &self.cur,
            attributes,
            self.depth,
        )));
        if let Some(cur_parent) = &self.cur {
            // 最新节点添加为 parent 的子节点
            cur_parent.borrow_mut().append_child(Rc::clone(&element));
            // cur parent 变更为 最新节点
            self.cur.replace(element);
        } else {
            // 第一个节点
            self.root.replace(Rc::clone(&element));
            self.cur.replace(element);
        }
    }
    fn fish_feed_element(&mut self) {
        self.depth -= 1;

        // 当前父节点指向上一层的节点
        let mut parent = None;
        if let Some(cur) = &self.cur {
            if let Some(p) = &cur.borrow().parent {
                if let v = p.upgrade() {
                    parent = v;
                }
            }
        }

        self.cur = parent;
    }
    /// 目前只有 w:t 类型会有
    fn feed_characters(&mut self, data: String) {
        if let Some(cur) = &self.cur {
            cur.borrow_mut().literal_text = Some(data);
        }
    }
}

fn print_elements(root: &Option<Rc<RefCell<Element>>>) {
    println!("{}", Element::display(root));
    if let Some(root_rc) = root {
        if root_rc.borrow().element_type.is_text() {
            let run_property = Element::find_run_property(&root);
            let color_val = Element::get_color(&run_property);
            let text = root_rc
                .borrow()
                .literal_text
                .as_ref()
                .map(|v| v.clone())
                .unwrap_or_default();
            println!("text: {}", text);
            if let Some(run_pr) = run_property {
                println!("----has run property---{:?}", color_val);
            }
        }
        for child in root_rc.borrow().children.iter() {
            print_elements(&Some(Rc::clone(child)));
        }
    }
}
