// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! DOS filename expressions used by the virtual Windows filesystem.

#[derive(Clone, Copy, PartialEq, Eq)]
enum Token { Literal(u16), Star, DosStar, DosQuestion, DosDot }

fn translate(pattern: &str) -> Vec<Token> {
    if matches!(pattern,""|"*"|"*.*") {return vec![Token::Star];}
    let units:Vec<_>=pattern.to_lowercase().encode_utf16().collect();
    let mut result=Vec::new();
    for (at,&unit) in units.iter().enumerate() {
        match unit {
            42=>result.push(Token::Star),
            63=>result.push(Token::DosQuestion),
            46 if at+1==units.len() && result.last()==Some(&Token::Star)=>{
                result.pop();result.push(Token::DosStar);
            },
            46 if units.get(at+1).is_some_and(|next|matches!(next,42|63))=>result.push(Token::DosDot),
            _=>result.push(Token::Literal(unit)),
        }
    }
    result
}

pub(super) fn matches(name: &str, pattern: &str) -> bool {
    let name:Vec<_>=name.to_lowercase().encode_utf16().collect();
    let expression=translate(pattern);
    let dot_limit=name.iter().rposition(|unit|*unit==46).map_or(name.len(),|at|at+1);
    let mut memo=vec![None;(expression.len()+1)*(name.len()+1)];
    fn walk(expression:&[Token],name:&[u16],at:usize,pos:usize,dot_limit:usize,memo:&mut [Option<bool>])->bool {
        let index=at*(name.len()+1)+pos;
        if let Some(value)=memo[index] {return value;}
        let value=match expression.get(at) {
            None=>pos==name.len(),
            Some(Token::Literal(unit))=>name.get(pos)==Some(unit) && walk(expression,name,at+1,pos+1,dot_limit,memo),
            Some(Token::Star)=>walk(expression,name,at+1,pos,dot_limit,memo)
                || pos<name.len() && walk(expression,name,at,pos+1,dot_limit,memo),
            Some(Token::DosStar)=>walk(expression,name,at+1,pos,dot_limit,memo)
                || pos<dot_limit && walk(expression,name,at,pos+1,dot_limit,memo),
            Some(Token::DosQuestion) if pos==name.len() || name.get(pos)==Some(&46)=>{
                let mut next=at+1;
                while expression.get(next)==Some(&Token::DosQuestion) {next+=1;}
                walk(expression,name,next,pos,dot_limit,memo)
            },
            Some(Token::DosQuestion)=>walk(expression,name,at+1,pos+1,dot_limit,memo),
            Some(Token::DosDot) if pos==name.len()=>walk(expression,name,at+1,pos,dot_limit,memo),
            Some(Token::DosDot)=>name.get(pos)==Some(&46) && walk(expression,name,at+1,pos+1,dot_limit,memo),
        };
        memo[index]=Some(value);value
    }
    walk(&expression,&name,0,0,dot_limit,&mut memo)
}

#[cfg(test)]
mod tests {
    use super::matches;
    #[test]
    fn native_dos_patterns_include_extensionless_names_and_optional_questions() {
        for (name,pattern) in [("plain","*.*"),("plain","plain.*"),("a.txt","a?.txt"),("ab.txt","a?.txt"),("a.txt","??.txt"),("UPPER.TXT","*.txt"),("name.a.b","name.*"),("plain","*.")] {
            assert!(matches(name,pattern),"{name} / {pattern}");
        }
        for (name,pattern) in [("a.txt","*."),("ab.txt","?.txt"),("a.txt","a?"),("name.a.b","?.txt")] {
            assert!(!matches(name,pattern),"{name} / {pattern}");
        }
    }
}
