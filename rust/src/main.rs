use saphyr_parser::{Event, Parser, ScalarStyle};

fn main() {
    let src = "a: 1\nb: [x, y]\nc: { k: v }\nnote: |\n  hello\n  world\nnum: 0.30\nd: 1904\nq: \"tru:e\"\ns: 'a''b'\ntrue_v: true\nnull_v: ~\nempty_v:\nlist:\n  - { name: R, value: '{\"AU\",\"NZ\"}' }\n";
    for item in Parser::new_from_str(src) {
        match item {
            Ok((ev, _span)) => match ev {
                Event::Scalar(v, style, anchor, tag) => {
                    let st = match style {
                        ScalarStyle::Plain => "plain",
                        ScalarStyle::SingleQuoted => "sq",
                        ScalarStyle::DoubleQuoted => "dq",
                        ScalarStyle::Literal => "lit",
                        ScalarStyle::Folded => "fold",
                    };
                    println!("Scalar[{st} a={anchor} tag={}]: {:?}", tag.is_some(), v);
                }
                Event::StreamEnd => {
                    println!("StreamEnd");
                    break;
                }
                other => println!("{other:?}"),
            },
            Err(e) => {
                println!("ERR {e:?}");
                break;
            }
        }
    }
}
