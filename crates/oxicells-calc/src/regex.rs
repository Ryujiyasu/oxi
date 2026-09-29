// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! A small backtracking regular-expression engine, enough for Excel's REGEX
//! functions and VBScript's RegExp object: literals and escapes, `.`,
//! classes with ranges and `\d \w \s` (and their negations), anchors `^ $
//! \b \B`, capturing and non-capturing groups, alternation, greedy and lazy
//! quantifiers `* + ? {n} {n,} {n,m}`, backreferences and lookahead. It
//! works on chars and leftmost-first, as Perl-style engines do.

#[derive(Debug, Clone)]
enum Node {
    Char(char),
    Any,
    Class { items: Vec<ClassItem>, negated: bool },
    Start,
    End,
    WordBoundary(bool),
    Group(Box<Node>, Option<usize>),
    Concat(Vec<Node>),
    Alternate(Vec<Node>),
    Repeat { node: Box<Node>, min: usize, max: Option<usize>, greedy: bool },
    BackReference(usize),
    LookAhead(Box<Node>, bool),
    LookBehind(Box<Node>, bool),
}

#[derive(Debug, Clone)]
enum ClassItem {
    Range(char, char),
    Digit(bool),
    Word(bool),
    Space(bool),
}

/// A compiled pattern.
#[derive(Debug, Clone)]
pub struct Regex {
    node: Node,
    groups: usize,
    ignore_case: bool,
    multiline: bool,
    /// How much work the current search has done, so that a pattern that
    /// backtracks without end gives up rather than hanging.
    steps: std::cell::Cell<usize>,
}

/// Where each group matched, the whole match first; None for a group that
/// took no part.
pub type Captures = Vec<Option<(usize, usize)>>;

struct Parser<'a> {
    chars: &'a [char],
    at: usize,
    groups: usize,
}

impl Parser<'_> {
    fn peek(&self) -> Option<char> {
        self.chars.get(self.at).copied()
    }

    fn alternation(&mut self) -> Result<Node, String> {
        let mut branches = vec![self.sequence()?];
        while self.peek() == Some('|') {
            self.at += 1;
            branches.push(self.sequence()?);
        }
        Ok(if branches.len() == 1 { branches.pop().unwrap_or(Node::Concat(Vec::new())) } else { Node::Alternate(branches) })
    }

    fn sequence(&mut self) -> Result<Node, String> {
        let mut items = Vec::new();
        while let Some(c) = self.peek() {
            if c == '|' || c == ')' {
                break;
            }
            let atom = self.atom()?;
            let atom = self.quantified(atom)?;
            items.push(atom);
        }
        Ok(Node::Concat(items))
    }

    fn number(&mut self) -> Option<usize> {
        let start = self.at;
        while matches!(self.peek(), Some(c) if c.is_ascii_digit()) {
            self.at += 1;
        }
        if self.at == start {
            return None;
        }
        self.chars[start..self.at].iter().collect::<String>().parse().ok()
    }

    fn quantified(&mut self, atom: Node) -> Result<Node, String> {
        let (min, max) = match self.peek() {
            Some('*') => {
                self.at += 1;
                (0, None)
            }
            Some('+') => {
                self.at += 1;
                (1, None)
            }
            Some('?') => {
                self.at += 1;
                (0, Some(1))
            }
            Some('{') => {
                let saved = self.at;
                self.at += 1;
                match self.number() {
                    Some(low) => {
                        let high = if self.peek() == Some(',') {
                            self.at += 1;
                            self.number()
                        } else {
                            Some(low)
                        };
                        if self.peek() != Some('}') {
                            self.at = saved;
                            return Ok(atom);
                        }
                        self.at += 1;
                        (low, high)
                    }
                    None => {
                        self.at = saved;
                        return Ok(atom);
                    }
                }
            }
            _ => return Ok(atom),
        };
        if matches!(atom, Node::Start | Node::End | Node::WordBoundary(_)) {
            return Err("nothing to repeat".to_string());
        }
        let greedy = if self.peek() == Some('?') {
            self.at += 1;
            false
        } else {
            if self.peek() == Some('+') {
                self.at += 1;
            }
            true
        };
        Ok(Node::Repeat { node: Box::new(atom), min, max, greedy })
    }

    fn escape_class(&mut self, c: char) -> Option<ClassItem> {
        Some(match c {
            'd' => ClassItem::Digit(true),
            'D' => ClassItem::Digit(false),
            'w' => ClassItem::Word(true),
            'W' => ClassItem::Word(false),
            's' => ClassItem::Space(true),
            'S' => ClassItem::Space(false),
            _ => return None,
        })
    }

    fn escaped_char(&mut self, c: char) -> Result<char, String> {
        Ok(match c {
            't' => '\t',
            'n' => '\n',
            'r' => '\r',
            'f' => '\u{0C}',
            'v' => '\u{0B}',
            '0' => '\0',
            'x' => {
                let hex: String = self.chars.get(self.at..self.at + 2).unwrap_or(&[]).iter().collect();
                self.at += 2;
                char::from_u32(u32::from_str_radix(&hex, 16).map_err(|_| "bad \\x")?).ok_or("bad \\x")?
            }
            'u' => {
                let hex: String = self.chars.get(self.at..self.at + 4).unwrap_or(&[]).iter().collect();
                self.at += 4;
                char::from_u32(u32::from_str_radix(&hex, 16).map_err(|_| "bad \\u")?).ok_or("bad \\u")?
            }
            other => other,
        })
    }

    fn atom(&mut self) -> Result<Node, String> {
        let c = self.peek().ok_or("unexpected end")?;
        self.at += 1;
        Ok(match c {
            '.' => Node::Any,
            '^' => Node::Start,
            '$' => Node::End,
            '(' => {
                let (capture, look) = if self.peek() == Some('?') {
                    match self.chars.get(self.at + 1) {
                        Some(':') => {
                            self.at += 2;
                            (None, None)
                        }
                        Some('=') => {
                            self.at += 2;
                            (None, Some(true))
                        }
                        Some('!') => {
                            self.at += 2;
                            (None, Some(false))
                        }
                        Some('<') if matches!(self.chars.get(self.at + 2), Some('=') | Some('!')) => {
                            let positive = self.chars[self.at + 2] == '=';
                            self.at += 3;
                            let inner = self.alternation()?;
                            if self.peek() != Some(')') {
                                return Err("missing )".to_string());
                            }
                            self.at += 1;
                            return Ok(Node::LookBehind(Box::new(inner), positive));
                        }
                        // A named group, (?<name>...) or (?P<name>...),
                        // is numbered with the rest.
                        Some('<') | Some('P') => {
                            while self.peek().is_some_and(|c| c != '>') {
                                self.at += 1;
                            }
                            self.at += 1;
                            self.groups += 1;
                            (Some(self.groups), None)
                        }
                        _ => return Err("unsupported group".to_string()),
                    }
                } else {
                    self.groups += 1;
                    (Some(self.groups), None)
                };
                let inner = self.alternation()?;
                if self.peek() != Some(')') {
                    return Err("missing )".to_string());
                }
                self.at += 1;
                match look {
                    Some(positive) => Node::LookAhead(Box::new(inner), positive),
                    None => Node::Group(Box::new(inner), capture),
                }
            }
            ')' => return Err("unmatched )".to_string()),
            '[' => self.class()?,
            '\\' => {
                let e = self.peek().ok_or("trailing \\")?;
                self.at += 1;
                if let Some(item) = self.escape_class(e) {
                    Node::Class { items: vec![item], negated: false }
                } else if e == 'b' || e == 'B' {
                    Node::WordBoundary(e == 'b')
                } else if e.is_ascii_digit() && e != '0' {
                    let mut n = e.to_digit(10).unwrap_or(0) as usize;
                    while let Some(d) = self.peek().and_then(|d| d.to_digit(10)) {
                        if n * 10 + d as usize > self.groups.max(9) {
                            break;
                        }
                        n = n * 10 + d as usize;
                        self.at += 1;
                    }
                    Node::BackReference(n)
                } else {
                    Node::Char(self.escaped_char(e)?)
                }
            }
            '*' | '+' | '?' => return Err("nothing to repeat".to_string()),
            other => Node::Char(other),
        })
    }

    fn class(&mut self) -> Result<Node, String> {
        let negated = if self.peek() == Some('^') {
            self.at += 1;
            true
        } else {
            false
        };
        let mut items = Vec::new();
        let mut first = true;
        loop {
            let c = self.peek().ok_or("missing ]")?;
            self.at += 1;
            if c == ']' && !first {
                break;
            }
            first = false;
            let low = if c == '\\' {
                let e = self.peek().ok_or("trailing \\")?;
                self.at += 1;
                if let Some(item) = self.escape_class(e) {
                    items.push(item);
                    continue;
                }
                if e == 'b' {
                    '\u{08}'
                } else {
                    self.escaped_char(e)?
                }
            } else {
                c
            };
            if self.peek() == Some('-') && self.chars.get(self.at + 1).is_some_and(|n| *n != ']') {
                self.at += 1;
                let h = self.peek().ok_or("missing ]")?;
                self.at += 1;
                let high = if h == '\\' {
                    let e = self.peek().ok_or("trailing \\")?;
                    self.at += 1;
                    self.escaped_char(e)?
                } else {
                    h
                };
                if high < low {
                    return Err("range out of order".to_string());
                }
                items.push(ClassItem::Range(low, high));
            } else {
                items.push(ClassItem::Range(low, low));
            }
        }
        Ok(Node::Class { items, negated })
    }
}

fn is_word(c: char) -> bool {
    c.is_alphanumeric() || c == '_'
}

fn fold(c: char) -> char {
    c.to_lowercase().next().unwrap_or(c)
}

impl Regex {
    pub fn new(pattern: &str, ignore_case: bool, multiline: bool) -> Result<Regex, String> {
        let chars: Vec<char> = pattern.chars().collect();
        let mut parser = Parser { chars: &chars, at: 0, groups: 0 };
        let node = parser.alternation()?;
        if parser.at < chars.len() {
            return Err("unmatched )".to_string());
        }
        Ok(Regex { node, groups: parser.groups, ignore_case, multiline, steps: std::cell::Cell::new(0) })
    }

    /// How many capturing groups the pattern has.
    pub fn group_count(&self) -> usize {
        self.groups
    }

    /// The first match starting at or after `start`.
    pub fn find_at(&self, text: &[char], start: usize) -> Option<Captures> {
        for from in start..=text.len() {
            let mut caps: Captures = vec![None; self.groups + 1];
            self.steps.set(0);
            let mut end_at = None;
            if self.walk(&self.node, text, from, &mut caps, &mut |at, _| {
                end_at = Some(at);
                true
            }) {
                caps[0] = Some((from, end_at.unwrap_or(from)));
                return Some(caps);
            }
        }
        None
    }

    /// Every match, left to right, an empty match moving on one character.
    pub fn find_all(&self, text: &[char]) -> Vec<Captures> {
        let mut out = Vec::new();
        let mut at = 0;
        while at <= text.len() {
            match self.find_at(text, at) {
                Some(caps) => {
                    let (s, e) = caps[0].unwrap_or((at, at));
                    at = if e == s { e + 1 } else { e };
                    out.push(caps);
                }
                None => break,
            }
        }
        out
    }

    fn class_matches(&self, items: &[ClassItem], negated: bool, c: char) -> bool {
        let hit = items.iter().any(|item| match item {
            ClassItem::Range(low, high) => {
                (*low..=*high).contains(&c)
                    || (self.ignore_case && {
                        let (l, u) = (fold(c), c.to_uppercase().next().unwrap_or(c));
                        (*low..=*high).contains(&l) || (*low..=*high).contains(&u)
                    })
            }
            ClassItem::Digit(yes) => c.is_ascii_digit() == *yes,
            ClassItem::Word(yes) => is_word(c) == *yes,
            ClassItem::Space(yes) => c.is_whitespace() == *yes,
        });
        hit != negated
    }

    /// Match `node` at `at`, then hand the end to `next`; true once `next`
    /// accepts. Captures are set on the way and undone on backtracking.
    fn walk(
        &self,
        node: &Node,
        text: &[char],
        at: usize,
        caps: &mut Captures,
        next: &mut dyn FnMut(usize, &mut Captures) -> bool,
    ) -> bool {
        self.steps.set(self.steps.get() + 1);
        if self.steps.get() > 5_000_000 {
            return false;
        }
        match node {
            Node::Char(c) => match text.get(at) {
                Some(t) if *t == *c || (self.ignore_case && fold(*t) == fold(*c)) => next(at + 1, caps),
                _ => false,
            },
            Node::Any => match text.get(at) {
                Some(t) if *t != '\n' => next(at + 1, caps),
                _ => false,
            },
            Node::Class { items, negated } => match text.get(at) {
                Some(t) if self.class_matches(items, *negated, *t) => next(at + 1, caps),
                _ => false,
            },
            Node::Start => {
                let ok = at == 0 || (self.multiline && text[at - 1] == '\n');
                ok && next(at, caps)
            }
            Node::End => {
                let ok = at == text.len() || (self.multiline && text[at] == '\n');
                ok && next(at, caps)
            }
            Node::WordBoundary(want) => {
                let before = at > 0 && is_word(text[at - 1]);
                let after = at < text.len() && is_word(text[at]);
                ((before != after) == *want) && next(at, caps)
            }
            Node::Group(inner, index) => match index {
                None => self.walk(inner, text, at, caps, next),
                Some(i) => {
                    let i = *i;
                    let saved = caps[i];
                    let start = at;
                    let ok = self.walk(inner, text, at, caps, &mut |end, caps: &mut Captures| {
                        let before = caps[i];
                        caps[i] = Some((start, end));
                        if next(end, caps) {
                            return true;
                        }
                        caps[i] = before;
                        false
                    });
                    if !ok {
                        caps[i] = saved;
                    }
                    ok
                }
            },
            Node::Concat(items) => self.sequence(items, text, at, caps, next),
            Node::Alternate(branches) => {
                for branch in branches {
                    if self.walk(branch, text, at, caps, next) {
                        return true;
                    }
                }
                false
            }
            Node::Repeat { node, min, max, greedy } => self.repeat(node, *min, *max, *greedy, 0, text, at, caps, next),
            Node::BackReference(i) => {
                let Some(Some((s, e))) = caps.get(*i).copied() else {
                    return next(at, caps);
                };
                let length = e - s;
                if at + length > text.len() {
                    return false;
                }
                let same = (0..length).all(|k| {
                    let (a, b) = (text[s + k], text[at + k]);
                    a == b || (self.ignore_case && fold(a) == fold(b))
                });
                same && next(at + length, caps)
            }
            Node::LookBehind(inner, positive) => {
                let mut probe = caps.clone();
                let found = (0..=at).rev().any(|from| self.walk(inner, text, from, &mut probe, &mut |end, _| end == at));
                if found == *positive {
                    if *positive {
                        *caps = probe;
                    }
                    next(at, caps)
                } else {
                    false
                }
            }
            Node::LookAhead(inner, positive) => {
                let mut probe = caps.clone();
                let found = self.walk(inner, text, at, &mut probe, &mut |_, _| true);
                if found == *positive {
                    if *positive {
                        *caps = probe;
                    }
                    next(at, caps)
                } else {
                    false
                }
            }
        }
    }

    fn sequence(
        &self,
        items: &[Node],
        text: &[char],
        at: usize,
        caps: &mut Captures,
        next: &mut dyn FnMut(usize, &mut Captures) -> bool,
    ) -> bool {
        match items.split_first() {
            None => next(at, caps),
            Some((first, rest)) => {
                self.walk(first, text, at, caps, &mut |end, caps: &mut Captures| self.sequence(rest, text, end, caps, next))
            }
        }
    }

    #[allow(clippy::too_many_arguments)]
    fn repeat(
        &self,
        node: &Node,
        min: usize,
        max: Option<usize>,
        greedy: bool,
        done: usize,
        text: &[char],
        at: usize,
        caps: &mut Captures,
        next: &mut dyn FnMut(usize, &mut Captures) -> bool,
    ) -> bool {
        let can_more = max.is_none_or(|m| done < m);
        let more = |caps: &mut Captures, next: &mut dyn FnMut(usize, &mut Captures) -> bool| {
            can_more
                && self.walk(node, text, at, caps, &mut |end, caps: &mut Captures| {
                    // An empty pass would go round for ever.
                    if end == at && done >= min {
                        return false;
                    }
                    self.repeat(node, min, max, greedy, done + 1, text, end, caps, next)
                })
        };
        if done < min {
            return more(caps, next);
        }
        if greedy {
            more(caps, next) || next(at, caps)
        } else {
            next(at, caps) || more(caps, next)
        }
    }
}

/// The replacement text for one match: `$n` and `${n}` are group n (`$0`
/// the whole match), `$$` a dollar; anything else is itself.
pub fn expand(replacement: &str, text: &[char], caps: &Captures) -> String {
    let r: Vec<char> = replacement.chars().collect();
    let group = |n: usize, out: &mut String| {
        if let Some(Some((s, e))) = caps.get(n) {
            out.extend(&text[*s..*e]);
        }
    };
    let mut out = String::new();
    let mut i = 0;
    while i < r.len() {
        if r[i] == '$' && i + 1 < r.len() {
            if r[i + 1] == '$' {
                out.push('$');
                i += 2;
                continue;
            }
            if r[i + 1] == '{' {
                if let Some(close) = r[i + 2..].iter().position(|c| *c == '}') {
                    let name: String = r[i + 2..i + 2 + close].iter().collect();
                    if let Ok(n) = name.parse::<usize>() {
                        group(n, &mut out);
                        i += close + 3;
                        continue;
                    }
                }
            }
            if r[i + 1].is_ascii_digit() {
                let mut j = i + 1;
                let mut n = 0usize;
                while j < r.len() && r[j].is_ascii_digit() && (j == i + 1 || n * 10 + (r[j] as usize - 48) < caps.len()) {
                    n = n * 10 + (r[j] as usize - 48);
                    j += 1;
                }
                group(n, &mut out);
                i = j;
                continue;
            }
        }
        out.push(r[i]);
        i += 1;
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn first(p: &str, t: &str) -> Option<String> {
        let re = Regex::new(p, false, false).ok()?;
        let chars: Vec<char> = t.chars().collect();
        re.find_at(&chars, 0).map(|c| {
            let (s, e) = c[0].unwrap();
            chars[s..e].iter().collect()
        })
    }

    #[test]
    fn matches_the_usual_forms() {
        assert_eq!(first("[0-9]+", "abc123def"), Some("123".into()));
        assert_eq!(first("a.c", "xxabcx"), Some("abc".into()));
        assert_eq!(first("(ab)+", "ababab"), Some("ababab".into()));
        assert_eq!(first("a+?", "aaa"), Some("a".into()));
        assert_eq!(first("^b", "ab"), None);
        assert_eq!(first("\\bcat\\b", "a cat!"), Some("cat".into()));
        assert_eq!(first("(\\w)\\1", "abccd"), Some("cc".into()));
        assert_eq!(first("x(?=y)", "xzxy"), Some("x".into()));
        assert_eq!(first("colou?r|gray", "the gray"), Some("gray".into()));
        assert_eq!(first("\\d{2,3}", "a1234"), Some("123".into()));
        assert_eq!(first("[^a-c]+", "abcxyz"), Some("xyz".into()));
        assert_eq!(first(r"(?<=y=)\d+","x=10,y=20"), Some("20".into()));
        assert_eq!(first("(?<name>a)b", "xab"), Some("ab".into()));
    }
}
