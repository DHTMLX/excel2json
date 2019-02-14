pub struct Range {
    pub first: (u32, u32),
    pub last: (u32, u32)
}

impl Range {
    pub fn new(dimension: String) -> Range {
        let split = dimension.split(":");
        let cells: Vec<String> = split.take(2).map(String::from).collect();

        let first = cell_index_to_offsets(cells[0].clone());
        let last = cell_index_to_offsets(cells[1].clone());

        Range { first, last }
    }
    pub fn get_max_offsets(&self) -> (u32, u32) {
        (self.last.0 + 1, self.last.1 + 1)
    }
}


pub fn cell_index_to_offsets(s: String) -> (u32, u32) {
    let mut alpha = vec!();
    let mut number_part: u32 = 0;
    for byte in s.into_bytes() {
        if byte >= 65 && byte <= 90 { // A - Z
            alpha.push(byte as u32 - 64);
        } else { // 0 - 9
            number_part = number_part * 10 + byte as u32 - 48;
        }
    }
    let len = alpha.len() as u32;
    let alpha_part = alpha.iter().enumerate().fold(0, |s, (index, v)| s + 26u32.pow(len - index as u32 - 1) * v);
    (alpha_part - 1, number_part - 1)
}