const THIN_BORDER: &str = "0.5px";
const MEDIUM_BORDER: &str = "1px";
const THICK_BORDER: &str = "2px";

pub enum BorderPosition {
    Left,
    Right,
    Top,
    Bottom,
}
enum BorderStyle {
    Medium,
    Dotted,
    Thick,
    Thin,
    Hair,
    Double,
}

pub struct Border {
    position: BorderPosition,
    style: Option<BorderStyle>,
    color: String,
}

impl Border {
    pub fn new(position: BorderPosition) -> Border {
        Border { position, color: String::from("#000000"), style: None }
    }
    pub fn set_style(&mut self, style: String) {
        if &style == "thin" {
            self.style = Some(BorderStyle::Thin);
        } else if &style == "double" {
            self.style = Some(BorderStyle::Double);
        } else if &style == "thick" {
            self.style = Some(BorderStyle::Thick);
        } else if &style == "dotted" {
            self.style = Some(BorderStyle::Dotted);
        } else if &style == "hair" {
            self.style = Some(BorderStyle::Hair);
        } else if &style == "medium" {
            self.style = Some(BorderStyle::Medium);
        }
    }
    pub fn set_color(&mut self, color: String) {
        self.color = color;
    }
    pub fn get_computed_style(self) -> (String, String) {
        if self.style.is_none() {
            return (String::from(""), String::from(""))
        }

        let position = match self.position {
            BorderPosition::Left => String::from("borderLeft"),
            BorderPosition::Right => String::from("borderRight"),
            BorderPosition::Top => String::from("borderTop"),
            BorderPosition::Bottom => String::from("borderBottom")
        };

        let style = match self.style {
            Some(BorderStyle::Dotted) => "dotted",
            Some(BorderStyle::Double) => "double",
            Some(BorderStyle::Hair) => "dashed",
            Some(BorderStyle::Medium) | Some(BorderStyle::Thin) | Some(BorderStyle::Thick) => "solid",
            None => "",
        };

        let size = match self.style {
            Some(BorderStyle::Thin) => THIN_BORDER,
            Some(BorderStyle::Thick) => THICK_BORDER,
            _ => MEDIUM_BORDER,
        };

        (position, format!("{} {} {}", size, style, self.color))
    }
}


#[test] 
fn test_border() {
    let mut b = Border::new(BorderPosition::Top);
    b.set_style(String::from("thin"));
    b.set_color(String::from("#DF5K3FD"));
    let (_, val) = b.get_computed_style();
    assert_eq!(val, "0.5px solid #DF5K3FD");

    b = Border::new(BorderPosition::Top);
    b.set_style(String::from("thick"));
    b.set_color(String::from("#DF5K3FD"));
    let (_, val) = b.get_computed_style();
    assert_eq!(val, "2px solid #DF5K3FD");

    b = Border::new(BorderPosition::Top);
    b.set_style(String::from("medium"));
    let (_, val) = b.get_computed_style();
    assert_eq!(val, "1px solid #000000");

    b = Border::new(BorderPosition::Top);
    let (_, val) = b.get_computed_style();
    assert_eq!(val, "");
}   