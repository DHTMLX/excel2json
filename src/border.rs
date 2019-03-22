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
    style: BorderStyle,
    color: String
}

impl Border {
    pub fn new(position: BorderPosition) -> Border {
        Border { position, color: String::from("#D4D4D4"), style: BorderStyle::Medium }
    }
    pub fn set_style(&mut self, style: String) {
        if &style == "thin" {
            self.style = BorderStyle::Thin;
        } else if &style == "double" {
            self.style = BorderStyle::Double;
        } else if &style == "thick" {
            self.style = BorderStyle::Thick;
        } else if &style == "dotted" {
            self.style = BorderStyle::Dotted;
        } else if &style == "hair" {
            self.style = BorderStyle::Hair;
        }

    }
    pub fn set_color(&mut self, color: String) {
        self.color = color;
    }
    pub fn get_computed_style(self) -> (String, String) {
        let position = match self.position {
            BorderPosition::Left => String::from("borderLeft"),
            BorderPosition::Right => String::from("borderRight"),
            BorderPosition::Top => String::from("borderTop"),
            BorderPosition::Bottom => String::from("borderBottom")
        };

        let style = match self.style {
            BorderStyle::Medium | BorderStyle::Thin | BorderStyle::Thick => "solid",
            BorderStyle::Dotted => "dotted",
            BorderStyle::Double => "double",
            BorderStyle::Hair => "dashed",
        };

        let size = match self.style {
            BorderStyle::Thin => THIN_BORDER,
            BorderStyle::Thick => THICK_BORDER,
            _ => MEDIUM_BORDER
        };

        (position, format!("{} {} {}", size, style, self.color))
    }
}