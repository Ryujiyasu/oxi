// This Source Code Form is subject to the terms of the Mozilla Public
// License, v. 2.0. If a copy of the MPL was not distributed with this
// file, You can obtain one at https://mozilla.org/MPL/2.0/.

//! The files a macro opens with `Open`, kept in memory for the run.
//!
//! A browser has no disk to give a macro, so the files it writes live here
//! and are there to read back until the run ends -- which is what a macro
//! writing a report and reading it back, or keeping a scratch file, needs.
//! Paths are compared the way Windows compares them: case aside, `/` and
//! `\` alike. Text is kept as UTF-8, so a byte count (`LOF`, `FileLen`)
//! matches Excel's for ASCII and not for text its ANSI code page writes in
//! two bytes.

use std::collections::BTreeMap;
mod wildcards;

/// Process directory facts, independent of project variables and file handles.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct DirectoryState {
    pub current_drive: char,
    /// One absolute remembered directory per available drive.
    pub directories: Vec<String>,
}

/// Closed filesystem contents, independent of a project's variables, handles
/// and Dir iterator. The host can carry this across project replacement.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct FileSystemSnapshot {
    pub folders: Vec<String>,
    pub files: Vec<(String, Vec<u8>)>,
    /// Attribute overrides; omitted entries use archive32/file or directory16.
    pub attributes: Vec<(String, i64)>,
}

/// How a file was opened.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum OpenMode {
    Input,
    Output,
    Append,
    /// `For Binary`: bytes read and written at a position.
    Binary,
    /// `For Random Len = n`: records of n bytes.
    Random(usize),
}

/// One open file.
#[derive(Debug, Clone)]
pub struct Handle {
    pub path: String,
    pub mode: OpenMode,
    /// Where the next read starts, in bytes.
    pub position: usize,
    /// The column the next `Print` field starts at, counted from 0.
    pub column: usize,
    /// The access granted when the file opened, independent of later flags.
    pub can_read: bool,
    pub can_write: bool,
    pub share_read: bool,
    pub share_write: bool,
}

#[derive(Debug, Clone)]
struct RecordLock {
    owner: i64,
    path: String,
    start: u64,
    end: u64,
}

/// Why a file operation failed, as the VBA error number it raises.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FileError {
    /// 5, Invalid procedure call or argument.
    InvalidArgument = 5,
    /// 52, Bad file name or number.
    BadNumber = 52,
    /// 53, File not found.
    NotFound = 53,
    /// 55, File already open.
    AlreadyOpen = 55,
    /// 58, File already exists.
    AlreadyExists = 58,
    /// 62, Input past end of file.
    PastEnd = 62,
    /// 63, Bad record number.
    BadRecordNumber = 63,
    /// 68, Device unavailable.
    DeviceUnavailable = 68,
    /// 70, Permission denied.
    PermissionDenied = 70,
    /// 75, Path/File access error.
    PathFileAccess = 75,
    /// 76, Path not found.
    PathNotFound = 76,
}

impl FileError {
    pub fn number(self) -> i64 {
        self as i64
    }

    pub fn description(self) -> &'static str {
        match self {
            FileError::InvalidArgument => "Invalid procedure call or argument",
            FileError::BadNumber => "Bad file name or number",
            FileError::NotFound => "File not found",
            FileError::AlreadyOpen => "File already open",
            FileError::AlreadyExists => "File already exists",
            FileError::PastEnd => "Input past end of file",
            FileError::BadRecordNumber => "Bad record number",
            FileError::DeviceUnavailable => "Device unavailable",
            FileError::PermissionDenied => "Permission denied",
            FileError::PathFileAccess => "Path/File access error",
            FileError::PathNotFound => "Path not found",
        }
    }
}

/// The files of a run and the numbers they are open under.
#[derive(Debug, Clone)]
pub struct Files {
    contents: BTreeMap<String, (String, Vec<u8>)>,
    open: BTreeMap<i64, Handle>,
    /// What `Dir` with no arguments goes on to answer.
    listing: Option<Vec<String>>,
    /// Folders made with MkDir or CreateFolder, by key.
    folders: std::collections::BTreeSet<String>,
    current_drive: char,
    drive_directories: BTreeMap<char, String>,
    attributes: BTreeMap<String, i64>,
    volume_labels: BTreeMap<String,String>,
    record_locks: Vec<RecordLock>,
}

impl Default for Files {
    fn default() -> Self {
        Self {
            contents: BTreeMap::new(), open: BTreeMap::new(), listing: None,
            folders: Default::default(), current_drive: 'C',
            drive_directories: BTreeMap::from([('C', "C:\\".to_string())]),
            attributes: BTreeMap::new(),
            volume_labels: BTreeMap::new(),
            record_locks: Vec::new(),
        }
    }
}

/// The last part of a path.
fn file_name(path: &str) -> &str {
    path.rsplit(['\\', '/']).next().unwrap_or(path)
}

/// Whether a file name answers a `Dir` or `Kill` pattern of `*` and `?`.
fn matches_pattern(name: &str, pattern: &str) -> bool {
    wildcards::matches(name,pattern)
}

impl Files {
    /// Validate and normalize an absolute path supplied by a host snapshot.
    /// Relative mounts use the ordinary file-context API instead.
    pub fn normalize_snapshot_path(&self, path: &str) -> Result<String, FileError> {
        let bytes=path.as_bytes();
        let drive=bytes.len()>=3 && bytes[0].is_ascii_alphabetic() && bytes[1]==b':' && matches!(bytes[2],b'\\'|b'/');
        let unc=path.replace('/',"\\");
        let unc=unc.strip_prefix("\\\\").is_some_and(|tail| {
            let mut parts=tail.split('\\');
            matches!((parts.next(),parts.next()),(Some(server),Some(share)) if !server.is_empty() && !share.is_empty() && server!="." && server!=".." && share!="." && share!="..")
        });
        if !(drive||unc) || path.chars().enumerate().any(|(at,ch)| ch.is_control() || matches!(ch,'?'|'*'|'"'|'<'|'>'|'|') || ch==':' && !(drive && at==1)) {
            return Err(FileError::InvalidArgument);
        }
        Ok(self.absolute_path(path))
    }

    pub fn closed_snapshot(&self) -> FileSystemSnapshot {
        FileSystemSnapshot {
            folders:self.folders.iter().map(|path| if path.len()==2 && path.ends_with(':') {format!("{path}\\")} else {path.clone()}).collect(),
            files:self.contents.values().cloned().collect(),
            attributes:self.attributes.iter().map(|(path,value)| {
                let path=self.contents.get(path).map(|(path,_)|path.clone()).unwrap_or_else(||if path.len()==2 && path.ends_with(':') {format!("{path}\\")} else {path.clone()});
                (path,*value)
            }).collect(),
        }
    }

    /// Replace the closed namespace atomically. No VBA handles or enumeration
    /// cursors are restored; the process directory is a separate host fact.
    pub fn restore_closed_snapshot(&mut self, snapshot: &FileSystemSnapshot) -> Result<(), FileError> {
        let mut folders=std::collections::BTreeSet::new();
        let mut contents=BTreeMap::new();
        for path in &snapshot.folders {
            let path=self.normalize_snapshot_path(path)?;
            if !folders.insert(path.to_lowercase().trim_end_matches('\\').to_string()) {return Err(FileError::InvalidArgument);}
        }
        for (path,bytes) in &snapshot.files {
            let path=self.normalize_snapshot_path(path)?;
            let key=path.to_lowercase();
            if path.ends_with('\\') || folders.contains(&key) || contents.insert(key,(path,bytes.clone())).is_some() {return Err(FileError::InvalidArgument);}
        }
        let mut attributes=BTreeMap::new();
        for (path,value) in &snapshot.attributes {
            let path=self.normalize_snapshot_path(path)?;
            let key=path.to_lowercase().trim_end_matches('\\').to_string();
            let folder=folders.contains(&key);
            if !(folder || contents.contains_key(&key)) || *value<0 || value & !55!=0
                || (value & 16!=0)!=folder || attributes.insert(key,*value).is_some() {
                return Err(FileError::InvalidArgument);
            }
        }
        self.folders=folders;
        self.contents=contents;
        self.attributes=attributes;
        self.open.clear();
        self.record_locks.clear();
        self.listing=None;
        Ok(())
    }

    pub fn directory_state(&self) -> DirectoryState {
        DirectoryState {
            current_drive: self.current_drive,
            directories: ('A'..='Z').filter(|drive| self.drive_exists(*drive))
                .map(|drive| self.current_directory(Some(&drive.to_string())).expect("available drive")).collect(),
        }
    }

    /// Restore an explicitly supplied process environment atomically. Folder
    /// and file mounts remain separate; this does not create files or handles.
    pub fn restore_directory_state(&mut self, state: &DirectoryState) -> Result<(), FileError> {
        if !state.current_drive.is_ascii_alphabetic() { return Err(FileError::InvalidArgument); }
        let current=state.current_drive.to_ascii_uppercase();
        let mut directories=BTreeMap::new();
        for path in &state.directories {
            let bytes=path.as_bytes();
            if bytes.len()<3 || !bytes[0].is_ascii_alphabetic() || bytes[1]!=b':' || !matches!(bytes[2],b'\\'|b'/')
                || path.chars().enumerate().any(|(at,ch)| ch.is_control() || matches!(ch,'?'|'*'|'"'|'<'|'>'|'|') || ch==':' && at!=1) {
                return Err(FileError::InvalidArgument);
            }
            let drive=char::from(bytes[0]).to_ascii_uppercase();
            if directories.insert(drive,self.absolute_path(path)).is_some() { return Err(FileError::InvalidArgument); }
        }
        if !directories.contains_key(&current) { return Err(FileError::InvalidArgument); }
        self.current_drive=current;
        self.drive_directories=directories;
        Ok(())
    }

    fn drive_exists(&self, drive: char) -> bool {
        let prefix = format!("{}:", drive.to_ascii_lowercase());
        self.drive_directories.contains_key(&drive)
            || self.volume_labels.contains_key(&prefix)
            || self.folders.iter().any(|path| path.starts_with(&prefix))
            || self.contents.keys().any(|path| path.starts_with(&prefix))
    }

    /// Windows keeps one current directory for each drive. ChDir updates the
    /// addressed drive's directory; ChDrive chooses which drive is active.
    pub fn current_directory(&self, drive: Option<&str>) -> Result<String, FileError> {
        let drive = match drive.and_then(|value| value.chars().next()) {
            Some(value) if value.is_ascii_alphabetic() => value.to_ascii_uppercase(),
            Some(_) => return Err(FileError::InvalidArgument),
            None => self.current_drive,
        };
        if !self.drive_exists(drive) { return Err(FileError::DeviceUnavailable); }
        Ok(self.drive_directories.get(&drive).cloned().unwrap_or_else(|| format!("{drive}:\\")))
    }

    pub fn change_drive(&mut self, drive: &str) -> Result<(), FileError> {
        let Some(first) = drive.chars().next() else { return Ok(()); };
        if !first.is_ascii_alphabetic() { return Err(FileError::InvalidArgument); }
        let first = first.to_ascii_uppercase();
        if !self.drive_exists(first) { return Err(FileError::DeviceUnavailable); }
        self.current_drive = first;
        Ok(())
    }

    pub fn change_directory(&mut self, path: &str) -> Result<(), FileError> {
        if path.is_empty() { return Err(FileError::PathNotFound); }
        let path = self.absolute_path(path);
        if !self.folder_exists(&path) { return Err(FileError::PathNotFound); }
        if path.as_bytes().get(1) == Some(&b':') {
            let drive = char::from(path.as_bytes()[0]).to_ascii_uppercase();
            self.drive_directories.insert(drive, path);
        }
        Ok(())
    }

    /// Resolve against the active drive and its remembered directory, then
    /// collapse dot segments. Filename components follow Win32's handling of
    /// trailing spaces and periods; written download names remain separate.
    pub fn absolute_path(&self, path: &str) -> String {
        let path = path.replace('/', "\\");
        let absolute = if path.starts_with("\\\\") {
            path
        } else if path.as_bytes().get(1) == Some(&b':') && path.as_bytes()[0].is_ascii_alphabetic() {
            let drive = char::from(path.as_bytes()[0]).to_ascii_uppercase();
            let tail = &path[2..];
            if tail.starts_with('\\') { format!("{drive}:{tail}") }
            else {
                let base = self.drive_directories.get(&drive).cloned().unwrap_or_else(|| format!("{drive}:\\"));
                format!("{}\\{tail}", base.trim_end_matches('\\'))
            }
        } else if path.starts_with('\\') {
            format!("{}:{path}", self.current_drive)
        } else {
            let base = self.drive_directories.get(&self.current_drive).cloned().unwrap_or_else(|| format!("{}:\\", self.current_drive));
            format!("{}\\{path}", base.trim_end_matches('\\'))
        };
        let (prefix, tail, floor) = if let Some(tail) = absolute.strip_prefix("\\\\") {
            ("\\\\".to_string(), tail, 2)
        } else {
            (format!("{}\\", &absolute[..2]), &absolute[3..], 0)
        };
        let mut parts = Vec::new();
        for part in tail.split('\\') {
            match part {
                "" | "." => {},
                ".." => { if parts.len() > floor { parts.pop(); } },
                _ => { let part=part.trim_end_matches([' ', '.']); if !part.is_empty() { parts.push(part); } },
            }
        }
        format!("{prefix}{}", parts.join("\\"))
    }

    fn key(&self, path: &str) -> String {
        self.absolute_path(path.trim()).to_lowercase()
    }

    fn pattern_key(&self, path: &str) -> String {
        let path=path.replace('/',"\\");
        let start=path.rfind('\\').map_or(0,|at|at+1).max(if path.as_bytes().get(1)==Some(&b':') {2} else {0});
        let leaf=&path[start..];
        if leaf.contains(['*','?']) && leaf.ends_with('.') {
            let temporary=self.absolute_path(&format!("{}x",&path[..start]));
            let parent=temporary.rsplit_once('\\').map_or("",|(parent,_)|parent);
            format!("{parent}\\{leaf}").to_lowercase()
        } else {self.key(&path)}
    }

    fn volume_root(&self, path: &str) -> String {
        let path=self.absolute_path(path).to_lowercase();
        if let Some(tail)=path.strip_prefix("\\\\") {
            format!("\\\\{}",tail.split('\\').take(2).collect::<Vec<_>>().join("\\"))
        } else {path[..2].to_string()}
    }

    /// Host volume facts are separate from a closed namespace or cwd snapshot.
    pub fn volume_facts(&self) -> Vec<(String,String)> {
        self.volume_labels.iter().map(|(root,label)| {
            let root=if root.len()==2 && root.ends_with(':') {format!("{}\\",root.to_ascii_uppercase())} else {root.clone()};
            (root,label.clone())
        }).collect()
    }

    pub fn restore_volume_facts(&mut self, facts: &[(String,String)]) -> Result<(),FileError> {
        let mut volumes=BTreeMap::new();
        for (root,label) in facts {
            let path=self.normalize_snapshot_path(root)?;
            let root=self.volume_root(&path);
            if path.trim_end_matches('\\').to_lowercase()!=root || volumes.insert(root,label.clone()).is_some() {
                return Err(FileError::InvalidArgument);
            }
        }
        self.volume_labels=volumes;Ok(())
    }

    pub fn attributes(&self, path: &str) -> Result<i64, FileError> {
        let key=self.key(path).trim_end_matches('\\').to_string();
        let default=if self.contents.contains_key(&key) {32}
            else if self.folder_exists(&format!("{key}\\")) {16}
            else {return Err(FileError::NotFound);};
        Ok(self.attributes.get(&key).copied().unwrap_or(default))
    }

    /// SetAttr replaces the writable bits; the directory bit is structural.
    pub fn set_attributes(&mut self, path: &str, value: i64) -> Result<(), FileError> {
        if value<0 || value & !39!=0 {return Err(FileError::InvalidArgument);}
        let previous=self.attributes(path)?;
        let key=self.key(path).trim_end_matches('\\').to_string();
        if self.open.values().any(|handle|handle.path==key && handle.mode!=OpenMode::Input) {return Err(FileError::AlreadyOpen);}
        self.attributes.insert(key,value | (previous & 16));
        Ok(())
    }

    /// FSO masks unsupported bits and permits updates through open handles.
    pub fn set_fso_attributes(&mut self, path: &str, value: i64) -> Result<(), FileError> {
        let previous=self.attributes(path)?;
        let key=self.key(path).trim_end_matches('\\').to_string();
        self.attributes.insert(key,(value & 39) | (previous & 16));
        Ok(())
    }

    fn is_read_only_key(&self, key: &str) -> bool {
        self.attributes.get(key).copied().unwrap_or(0) & 1!=0
    }

    fn mark_written(&mut self, key: &str) {
        let attributes=self.attributes.get(key).copied().unwrap_or(32) | 32;
        self.attributes.insert(key.to_string(),attributes);
    }

    /// Publish complete binary contents to a closed file in an existing folder.
    /// Validation happens before replacement, so a refused write is atomic.
    pub fn replace_bytes(&mut self, path: &str, bytes: Vec<u8>) -> Result<(), FileError> {
        let name = self.key(path);
        if name.is_empty() || name.ends_with('\\') {
            return Err(FileError::BadNumber);
        }
        if self.open.values().any(|handle| handle.path == name) {
            return Err(FileError::AlreadyOpen);
        }
        if self.is_read_only_key(&name) {return Err(FileError::PermissionDenied);}
        if let Some((parent, _)) = name.rsplit_once('\\') {
            if !parent.is_empty() && !self.folder_exists(parent) {
                return Err(FileError::PathNotFound);
            }
        }
        self.mark_written(&name);
        self.contents.insert(name, (self.absolute_path(path.trim()), bytes));
        Ok(())
    }

    /// Read the bytes actually held by the VBA filesystem, including edits
    /// made through an open binary handle after a host produced the file.
    pub fn bytes(&self, path: &str) -> Result<&[u8], FileError> {
        self.contents.get(&self.key(path)).map(|(_, bytes)| bytes.as_slice()).ok_or(FileError::NotFound)
    }

    /// Absolute paths and complete contents of files created or modified since a snapshot.
    /// Renaming a newly produced file returns its final name; deleted files
    /// have no downloadable contents.
    pub fn changed_files(&self, before: &Self) -> Vec<(String, Vec<u8>)> {
        self.contents.iter().filter_map(|(key, (path, bytes))| {
            (before.contents.get(key).is_none_or(|(_, old)| old != bytes))
                .then(|| (path.clone(), bytes.clone()))
        }).collect()
    }

    pub fn deleted_files(&self, before: &Self) -> Vec<String> {
        before.contents.iter().filter(|(key, _)| !self.contents.contains_key(*key))
            .map(|(_, (path, _))| path.clone()).collect()
    }

    /// `FreeFile`: the lowest number not in use, from 1, or from 256 for
    /// `FreeFile(1)`.
    pub fn free_number(&self, upper_range: bool) -> i64 {
        let start = if upper_range { 256 } else { 1 };
        (start..start + 255).find(|number| !self.open.contains_key(number)).unwrap_or(0)
    }

    pub fn open(&mut self, path: &str, mode: OpenMode, number: i64) -> Result<(), FileError> {
        self.open_with_access(path,mode,number,None)
    }

    pub fn open_with_access(&mut self, path: &str, mode: OpenMode, number: i64, access: Option<crate::ast::FileAccess>) -> Result<(), FileError> {
        self.open_with_options(path,mode,number,access,None)
    }

    pub fn open_with_options(&mut self, path: &str, mode: OpenMode, number: i64, access: Option<crate::ast::FileAccess>, lock: Option<crate::ast::FileLock>) -> Result<(), FileError> {
        if !(1..=511).contains(&number) {
            return Err(FileError::BadNumber);
        }
        if self.open.contains_key(&number) {
            return Err(FileError::AlreadyOpen);
        }
        let name = self.key(path);
        if path.trim().is_empty() {
            return Err(FileError::BadNumber);
        }
        if self.open.iter().any(|(held_number,held)|*held_number<=511 && held.path==name &&
            (matches!(mode,OpenMode::Output|OpenMode::Append) || matches!(held.mode,OpenMode::Output|OpenMode::Append))) {
            return Err(FileError::AlreadyOpen);
        }
        let read_only=self.is_read_only_key(&name);
        let (can_read,can_write)=match access {
            Some(crate::ast::FileAccess::Read)=>(true,false),
            Some(crate::ast::FileAccess::Write)=>(false,true),
            Some(crate::ast::FileAccess::ReadWrite)=>(true,true),
            None=>match mode {
                OpenMode::Input=>(true,false),
                OpenMode::Output|OpenMode::Append=>(false,true),
                // VBA retries default binary/random access as read-only.
                OpenMode::Binary|OpenMode::Random(_)=>(true,!read_only),
            },
        };
        if read_only && can_write {return Err(FileError::PathFileAccess);}
        let (share_read,share_write)=match lock {
            None|Some(crate::ast::FileLock::Shared)=>(true,true),
            Some(crate::ast::FileLock::Read)=>(false,true),
            Some(crate::ast::FileLock::Write)=>(true,false),
            Some(crate::ast::FileLock::ReadWrite)=>(false,false),
        };
        if self.sharing_conflict(&name,can_read,can_write,share_read,share_write) {
            return Err(FileError::PermissionDenied);
        }
        match mode {
            OpenMode::Input if !self.contents.contains_key(&name) => return Err(FileError::NotFound),
            OpenMode::Output => {
                self.mark_written(&name);
                self.contents.insert(name.clone(), (self.absolute_path(path.trim()), Vec::new()));
            }
            OpenMode::Append | OpenMode::Binary | OpenMode::Random(_) => {
                let path=self.absolute_path(path.trim());
                self.contents.entry(name.clone()).or_insert_with(|| (path, Vec::new()));
            }
            OpenMode::Input => {}
        }
        let column = match mode {
            OpenMode::Append => {
                let bytes = &self.contents[&name].1;
                let last_line = bytes.iter().rposition(|byte| *byte == b'\n').map_or(0, |at| at + 1);
                bytes.len() - last_line
            }
            _ => 0,
        };
        self.open.insert(number, Handle { path: name, mode, position: 0, column,can_read,can_write,share_read,share_write });
        Ok(())
    }

    /// A number for a TextStream, from 1000 on, never one FreeFile hands out.
    pub fn open_stream(&mut self, path: &str, mode: OpenMode) -> Result<i64, FileError> {
        let number = (1000..).find(|number| !self.open.contains_key(number)).unwrap_or(1000);
        let name = self.key(path);
        if mode!=OpenMode::Input && self.is_read_only_key(&name) {return Err(FileError::PermissionDenied);}
        // Native FSO streams permit concurrent readers. ForWriting and
        // ForAppending deny writers, including writers that opened first.
        let can_read=mode==OpenMode::Input;let can_write=!can_read;
        let share_read=true;let share_write=can_read;
        if self.sharing_conflict(&name,can_read,can_write,share_read,share_write) {return Err(FileError::PermissionDenied);}
        match mode {
            OpenMode::Input if !self.contents.contains_key(&name) => return Err(FileError::NotFound),
            OpenMode::Output => {
                self.mark_written(&name);
                self.contents.insert(name.clone(), (self.absolute_path(path.trim()), Vec::new()));
            }
            _ => {
                let path=self.absolute_path(path.trim());
                self.contents.entry(name.clone()).or_insert_with(|| (path, Vec::new()));
            }
        }
        self.open.insert(number, Handle { path: name, mode, position: 0, column: 0,can_read,can_write,share_read,share_write });
        Ok(number)
    }

    fn sharing_conflict(&self, path: &str, can_read: bool, can_write: bool, share_read: bool, share_write: bool) -> bool {
        self.open.values().any(|held|held.path==path &&
            (can_read && !held.share_read || can_write && !held.share_write || held.can_read && !share_read || held.can_write && !share_write))
    }

    pub fn exists(&self, path: &str) -> bool {
        self.contents.contains_key(&self.key(path))
    }

    /// Whether the folder was made, or any kept file lies in it.
    pub fn folder_exists(&self, path: &str) -> bool {
        let folder = self.key(path);
        let folder = folder.trim_end_matches('\\');
        if folder.len()==2 && folder.as_bytes()[1]==b':' {
            return self.drive_exists(char::from(folder.as_bytes()[0]).to_ascii_uppercase());
        }
        self.folders.contains(folder)
            || self.folders.iter().any(|held| held.starts_with(&format!("{folder}\\")))
            || self.contents.keys().any(|held| held.rsplit_once('\\').is_some_and(|(parent, _)| parent == folder || parent.starts_with(&format!("{folder}\\"))))
    }

    /// Import a host directory fact together with its existing ancestors. This
    /// seeds the environment rather than performing VBA's one-level MkDir.
    pub fn mount_folder(&mut self, path: &str) -> Result<(), FileError> {
        if path.trim().is_empty() {return Err(FileError::InvalidArgument);}
        let path=self.normalize_snapshot_path(&self.absolute_path(path.trim()))?;
        let mut key=path.trim_end_matches('\\').to_lowercase();
        let floor=if let Some(tail)=key.strip_prefix("\\\\") {
            2+tail.split('\\').take(2).map(str::len).sum::<usize>()+1
        } else {2};
        let mut folders=Vec::new();
        loop {
            if self.contents.contains_key(&key) {return Err(FileError::AlreadyExists);}
            folders.push(key.clone());
            if key.len()<=floor {break;}
            let parent=key.rfind('\\').ok_or(FileError::InvalidArgument)?;
            key.truncate(parent);
        }
        self.folders.extend(folders);
        Ok(())
    }

    /// Create one directory, with an existing parent. The caller distinguishes
    /// VBA MkDir's error75 from FSO CreateFolder's error58 for an existing name.
    pub fn make_folder(&mut self, path: &str) -> Result<(), FileError> {
        if path.is_empty() {return Err(FileError::PathNotFound);}
        let folder = self.key(path).trim_end_matches('\\').to_string();
        if self.folder_exists(&format!("{folder}\\")) || self.contents.contains_key(&folder) {
            return Err(FileError::AlreadyExists);
        }
        let parent=folder.rsplit_once('\\').map(|(parent,_)|parent).ok_or(FileError::PathNotFound)?;
        if !self.folder_exists(&format!("{parent}\\")) {return Err(FileError::PathNotFound);}
        self.folders.insert(folder);
        Ok(())
    }

    /// RmDir refuses a nonempty or current directory without changing it. FSO
    /// deletion visits entries in name order and retains completed work when a
    /// later entry is locked; a caught error must expose those partial effects.
    pub fn remove_folder(&mut self, path: &str, with_contents: bool) -> Result<(), FileError> {
        if with_contents {return self.delete_folder_with_force(path,false);}
        if path.is_empty() {return Err(FileError::PathNotFound);}
        let folder = self.key(path).trim_end_matches('\\').to_string();
        if !self.folder_exists(&format!("{folder}\\")) {
            return Err(FileError::PathNotFound);
        }
        if self.is_read_only_key(&folder) {return Err(FileError::PathFileAccess);}
        let inside = format!("{folder}\\");
        let cwd=self.key(&self.current_directory(None)?).trim_end_matches('\\').to_string();
        if folder==cwd || cwd.starts_with(&inside) || folder.len()==2 && folder.ends_with(':')
            || self.contents.keys().any(|held|held.starts_with(&inside))
            || self.folders.iter().any(|held|held.starts_with(&inside)) {
            return Err(FileError::PathFileAccess);
        }
        self.folders.remove(&folder);
        self.attributes.remove(&folder);
        Ok(())
    }

    fn child_folders(&self, parent: &str) -> std::collections::BTreeSet<String> {
        let prefix=format!("{parent}\\");
        self.folders.iter().map(String::as_str)
            .chain(self.contents.keys().filter_map(|path|path.rsplit_once('\\').map(|(parent,_)|parent)))
            .filter_map(|path|path.strip_prefix(&prefix).and_then(|tail|tail.split('\\').next())
                .filter(|name|!name.is_empty()).map(|name|format!("{prefix}{name}"))).collect()
    }

    fn delete_directory_tree(&mut self, folder: &str, force: bool) -> Result<(), FileError> {
        let mut entries:BTreeMap<String,bool>=self.child_folders(folder).into_iter().map(|path|(path,true)).collect();
        for path in self.contents.keys().filter(|path|path.rsplit_once('\\').is_some_and(|(parent,_)|parent==folder)) {
            entries.insert(path.clone(),false);
        }
        for (path,directory) in entries {
            if directory {self.delete_directory_tree(&path,force)?;}
            else {
                if self.open.values().any(|handle|handle.path==path) || !force && self.is_read_only_key(&path) {return Err(FileError::PermissionDenied);}
                self.contents.remove(&path);
                self.attributes.remove(&path);
            }
        }
        let cwd=self.key(&self.current_directory(None)?).trim_end_matches('\\').to_string();
        if folder==cwd || cwd.starts_with(&format!("{folder}\\")) {return Err(FileError::PermissionDenied);}
        if !force && self.is_read_only_key(folder) {return Err(FileError::PermissionDenied);}
        self.folders.remove(folder);
        self.attributes.remove(folder);
        Ok(())
    }

    pub fn delete_folder_with_force(&mut self, path: &str, force: bool) -> Result<(), FileError> {
        if path.is_empty() {return Err(FileError::InvalidArgument);}
        let supplied=path.replace('/',"\\");
        let leaf=supplied.trim_end_matches('\\').rsplit('\\').next().unwrap_or_default();
        if matches!(leaf,"."|"..") {return Err(FileError::PathNotFound);}
        if supplied.ends_with('\\') {
            let absolute=supplied.starts_with("\\\\") || supplied.as_bytes().get(1)==Some(&b':') && supplied.as_bytes().get(2)==Some(&b'\\');
            return Err(if absolute {FileError::PathNotFound} else {FileError::InvalidArgument});
        }
        let pattern=self.key(path);
        let (parent,leaf)=pattern.rsplit_once('\\').ok_or(FileError::PathNotFound)?;
        let targets:Vec<_>=self.child_folders(parent).into_iter().filter(|folder|matches_pattern(file_name(folder),leaf)).collect();
        if targets.is_empty() {
            // FSO's folder matcher accepts a matching ordinary file but does
            // not remove it. A nonexistent name/pattern still raises76.
            return if self.contents.keys().any(|file|file.rsplit_once('\\').is_some_and(|(held,name)|held==parent && matches_pattern(name,leaf))) {
                Ok(())
            } else {Err(FileError::PathNotFound)};
        }
        for folder in targets {self.delete_directory_tree(&folder,force)?;}
        Ok(())
    }

    /// The resolved absolute paths of the files directly in a folder, in order.
    pub fn files_in(&self, path: &str) -> Vec<String> {
        let folder = self.key(path).trim_end_matches('\\').to_string();
        self.contents
            .iter()
            .filter(|(held, _)| held.rsplit_once('\\').is_some_and(|(parent, _)| parent == folder))
            .map(|(_, (written, _))| written.clone())
            .collect()
    }

    /// What is left to read, for ReadAll.
    pub fn read_rest(&mut self, number: i64) -> Result<String, FileError> {
        let (bytes, position) = self.reading(number)?;
        if *position >= bytes.len() {
            return Err(FileError::PastEnd);
        }
        let rest = String::from_utf8_lossy(&bytes[*position..]).into_owned();
        *position = bytes.len();
        Ok(rest)
    }

    /// `Close`: the numbers given, or every file when none is.
    pub fn close(&mut self, numbers: &[i64]) {
        if numbers.is_empty() {
            self.open.clear();
            self.record_locks.clear();
        } else {
            for number in numbers {
                self.open.remove(number);
            }
            self.record_locks.retain(|held|!numbers.contains(&held.owner));
        }
    }

    pub fn handle(&self, number: i64) -> Result<&Handle, FileError> {
        self.open.get(&number).ok_or(FileError::BadNumber)
    }

    pub(crate) fn restore_byte_position(&mut self, number: i64, path: &str, position: usize) {
        if let Some(handle)=self.open.get_mut(&number).filter(|handle|handle.path==path) {handle.position=position;}
    }

    pub fn lock_records(&mut self, number: i64, start: Option<i64>, end: Option<i64>, unlock: bool) -> Result<(),FileError> {
        let handle=self.handle(number)?.clone();
        let (start,end)=if start.is_none() && end.is_none() {(0,u64::MAX)} else {
            let start=start.unwrap_or(1);let end=end.unwrap_or(start);
            if start<1 || end<start {return Err(FileError::BadRecordNumber);}
            let unit=match handle.mode {OpenMode::Random(length)=>length.max(1) as u64,_=>1};
            ((start as u64-1).checked_mul(unit).ok_or(FileError::BadRecordNumber)?,(end as u64).checked_mul(unit).ok_or(FileError::BadRecordNumber)?)
        };
        if unlock {
            let index=self.record_locks.iter().position(|held|held.owner==number && held.path==handle.path && held.start==start && held.end==end).ok_or(FileError::PermissionDenied)?;
            self.record_locks.remove(index);
        } else {
            if self.record_locks.iter().any(|held|held.path==handle.path && held.start<end && start<held.end) {return Err(FileError::PermissionDenied);}
            self.record_locks.push(RecordLock {owner:number,path:handle.path.clone(),start,end});
        }
        Ok(())
    }

    fn check_record_access(&self, number: i64, handle: &Handle, count: usize) -> Result<(),FileError> {
        let count=match handle.mode {OpenMode::Random(length)=>count.max(length),_=>count};
        let start=handle.position as u64;let end=start.saturating_add(count as u64);
        if count!=0 && self.record_locks.iter().any(|held|held.path==handle.path && held.owner!=number && held.start<end && start<held.end) {
            Err(FileError::PermissionDenied)
        } else {Ok(())}
    }

    /// Write text to a file open for output, keeping track of its column.
    pub fn write(&mut self, number: i64, text: &str) -> Result<(), FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if handle.mode == OpenMode::Input {
            return Err(FileError::BadNumber);
        }
        if !handle.can_write {return Err(FileError::PathFileAccess);}
        let file = self.contents.get_mut(&handle.path).ok_or(FileError::BadNumber)?;
        file.1.extend_from_slice(text.as_bytes());
        match text.rfind('\n') {
            Some(at) => handle.column = text[at + 1..].chars().count(),
            None => handle.column += text.chars().count(),
        }
        let path=handle.path.clone();
        self.mark_written(&path);
        Ok(())
    }

    pub fn column(&self, number: i64) -> Result<usize, FileError> {
        Ok(self.handle(number)?.column)
    }

    /// `LOF`: the file's length in bytes.
    pub fn length_of_open(&self, number: i64) -> Result<usize, FileError> {
        let handle = self.handle(number)?;
        Ok(self.contents.get(&handle.path).map_or(0, |file| file.1.len()))
    }

    /// `EOF`: whether reading has reached the end.
    pub fn at_end(&self, number: i64) -> Result<bool, FileError> {
        let handle = self.handle(number)?;
        let length = self.contents.get(&handle.path).map_or(0, |file| file.1.len());
        Ok(handle.position >= length)
    }

    /// `Loc`: for a binary file the last byte read or written, for a
    /// random one the last record, and for a sequential one the bytes read
    /// so far over 128.
    pub fn location(&self, number: i64) -> Result<i64, FileError> {
        let handle = self.handle(number)?;
        Ok(match handle.mode {
            OpenMode::Binary => handle.position as i64,
            OpenMode::Random(length) => (handle.position / length.max(1)) as i64,
            _ => (handle.position as i64 + 127) / 128,
        })
    }

    /// `Seek(n)`: where the next read or write goes, from 1 -- a byte, or a
    /// record in a random file.
    pub fn next_position(&self, number: i64) -> Result<i64, FileError> {
        let handle = self.handle(number)?;
        Ok(match handle.mode {
            OpenMode::Random(length) => (handle.position / length.max(1)) as i64 + 1,
            _ => handle.position as i64 + 1,
        })
    }

    /// `Seek #n, position`, and the position a `Get` or `Put` names.
    pub fn move_to(&mut self, number: i64, position: i64) -> Result<(), FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if position < 1 {
            return Err(FileError::BadNumber);
        }
        handle.position = match handle.mode {
            OpenMode::Random(length) => (position as usize - 1) * length,
            _ => position as usize - 1,
        };
        Ok(())
    }

    /// The record length of a random file, or None.
    pub fn record_length(&self, number: i64) -> Result<Option<usize>, FileError> {
        Ok(match self.handle(number)?.mode {
            OpenMode::Random(length) => Some(length),
            _ => None,
        })
    }

    /// `Put`: bytes at the current position, the file growing (with
    /// noughts across any gap) to take them.
    pub fn put(&mut self, number: i64, bytes: &[u8]) -> Result<(), FileError> {
        let handle = self.handle(number)?;
        if !matches!(handle.mode, OpenMode::Binary | OpenMode::Random(_)) {
            return Err(FileError::BadNumber);
        }
        if !handle.can_write {return Err(FileError::PathFileAccess);}
        self.check_record_access(number,handle,bytes.len())?;
        let handle=self.open.get_mut(&number).expect("checked handle");
        let file = self.contents.get_mut(&handle.path).ok_or(FileError::BadNumber)?;
        let end = handle.position + bytes.len();
        if file.1.len() < end {
            file.1.resize(end, 0);
        }
        file.1[handle.position..end].copy_from_slice(bytes);
        handle.position = end;
        let path=handle.path.clone();
        self.mark_written(&path);
        Ok(())
    }

    /// After a whole `Get` or `Put` in a random file, the next record: one
    /// read or written short still takes its whole length.
    pub fn finish_record(&mut self, number: i64) {
        if let Some(handle) = self.open.get_mut(&number) {
            if let OpenMode::Random(length) = handle.mode {
                let length = length.max(1);
                handle.position = handle.position.div_ceil(length) * length;
            }
        }
    }

    /// `Get`: so many bytes from the current position, noughts past the end.
    pub fn get(&mut self, number: i64, count: usize) -> Result<Vec<u8>, FileError> {
        let handle = self.handle(number)?;
        if !matches!(handle.mode, OpenMode::Binary | OpenMode::Random(_)) {
            return Err(FileError::BadNumber);
        }
        if !handle.can_read {return Err(FileError::PathFileAccess);}
        self.check_record_access(number,handle,count)?;
        let handle=self.open.get_mut(&number).expect("checked handle");
        let file = self.contents.get(&handle.path).map(|file| file.1.as_slice()).unwrap_or(&[]);
        let mut read = vec![0u8; count];
        for (index, slot) in read.iter_mut().enumerate() {
            if let Some(byte) = file.get(handle.position + index) {
                *slot = *byte;
            }
        }
        handle.position += count;
        Ok(read)
    }

    fn reading(&mut self, number: i64) -> Result<(&[u8], &mut usize), FileError> {
        let handle = self.open.get_mut(&number).ok_or(FileError::BadNumber)?;
        if handle.mode != OpenMode::Input {
            return Err(FileError::BadNumber);
        }
        if !handle.can_read {return Err(FileError::PathFileAccess);}
        let bytes = self.contents.get(&handle.path).map(|file| file.1.as_slice()).unwrap_or(&[]);
        Ok((bytes, &mut handle.position))
    }

    /// `Line Input #`: the next line, without its line break.
    pub fn read_line(&mut self, number: i64) -> Result<String, FileError> {
        let (bytes, position) = self.reading(number)?;
        if *position >= bytes.len() {
            return Err(FileError::PastEnd);
        }
        let start = *position;
        let mut end = start;
        while end < bytes.len() && bytes[end] != b'\r' && bytes[end] != b'\n' {
            end += 1;
        }
        let line = String::from_utf8_lossy(&bytes[start..end]).into_owned();
        let mut next = end;
        if next < bytes.len() && bytes[next] == b'\r' {
            next += 1;
        }
        if next < bytes.len() && bytes[next] == b'\n' {
            next += 1;
        }
        *position = next;
        Ok(line)
    }

    /// `Input(n, #f)`: the next n characters, line breaks and all.
    pub fn read_characters(&mut self, number: i64, count: usize) -> Result<String, FileError> {
        let (bytes, position) = self.reading(number)?;
        if *position + count > bytes.len() {
            return Err(FileError::PastEnd);
        }
        let read = String::from_utf8_lossy(&bytes[*position..*position + count]).into_owned();
        *position += count;
        Ok(read)
    }

    /// `Input #`: the next field as written, and whether it was quoted.
    pub fn read_field(&mut self, number: i64) -> Result<(String, bool), FileError> {
        let (bytes, position) = self.reading(number)?;
        let mut at = *position;
        while at < bytes.len() && matches!(bytes[at], b' ' | b'\t') {
            at += 1;
        }
        if at >= bytes.len() {
            return Err(FileError::PastEnd);
        }
        let (field, quoted) = if bytes[at] == b'"' {
            let start = at + 1;
            let mut end = start;
            while end < bytes.len() && bytes[end] != b'"' {
                end += 1;
            }
            at = (end + 1).min(bytes.len());
            // Past the closing quote, up to the next delimiter.
            while at < bytes.len() && !matches!(bytes[at], b',' | b'\r' | b'\n') {
                at += 1;
            }
            (String::from_utf8_lossy(&bytes[start..end]).into_owned(), true)
        } else {
            let start = at;
            while at < bytes.len() && !matches!(bytes[at], b',' | b'\r' | b'\n') {
                at += 1;
            }
            (String::from_utf8_lossy(&bytes[start..at]).trim().to_string(), false)
        };
        // One delimiter goes with the field: a comma, or a line break.
        if at < bytes.len() && bytes[at] == b',' {
            at += 1;
        } else {
            if at < bytes.len() && bytes[at] == b'\r' {
                at += 1;
            }
            if at < bytes.len() && bytes[at] == b'\n' {
                at += 1;
            }
        }
        *position = at;
        Ok((field, quoted))
    }

    /// `FileLen`: a closed or open file's length.
    pub fn length(&self, path: &str) -> Result<usize, FileError> {
        self.contents.get(&self.key(path)).map(|file| file.1.len()).ok_or(FileError::NotFound)
    }

    fn names_matching(&self, pattern: &str) -> Vec<String> {
        let wanted = self.pattern_key(pattern);
        let (folder, leaf) = match wanted.rfind('\\') {
            Some(at) => (&wanted[..=at], &wanted[at + 1..]),
            None => ("", wanted.as_str()),
        };
        self.contents
            .iter()
            .filter(|(held, _)| {
                let (held_folder, held_leaf) = match held.rfind('\\') {
                    Some(at) => (&held[..=at], &held[at + 1..]),
                    None => ("", held.as_str()),
                };
                held_folder == folder && matches_pattern(held_leaf, leaf)
            })
            .map(|(_, (written, _))| file_name(written).to_string())
            .collect()
    }

    /// `Dir(pattern)`: the first file answering it, or "" -- and `Dir()`
    /// the next.
    pub fn dir(&mut self, pattern: Option<&str>) -> String {
        self.dir_with_attributes(pattern,0)
    }

    pub fn dir_with_attributes(&mut self, pattern: Option<&str>, attributes: i64) -> String {
        self.dir_checked(pattern,attributes).unwrap_or_default()
    }

    /// Dir distinguishes its terminal empty answer from a later invalid call.
    pub fn dir_checked(&mut self, pattern: Option<&str>, attributes: i64) -> Result<String,FileError> {
        if attributes<0 || attributes & !63!=0 {return Err(FileError::InvalidArgument);}
        if let Some(pattern) = pattern {
            if attributes==8 {
                let root=self.volume_root(pattern);
                let label=self.volume_labels.get(&root).cloned().unwrap_or_default();
                self.listing=Some(Vec::new());
                return Ok(label);
            }
            let pattern=if pattern.is_empty() || pattern.ends_with(['\\','/']) {format!("{pattern}*")} else {pattern.to_string()};
            let wanted=self.pattern_key(&pattern);
            let folder=wanted.rfind('\\').map_or("",|at|&wanted[..=at]);
            let mut entries:BTreeMap<String,(String,i64)>=self.names_matching(&pattern).into_iter().map(|name| {
                let key=format!("{folder}{}",name.to_lowercase());
                let held=self.attributes.get(&key).copied().unwrap_or(32);
                (name.to_lowercase(),(name,held))
            }).collect();
            if attributes & 16!=0 {
                let parent=folder.trim_end_matches('\\');
                if self.folder_exists(&format!("{parent}\\")) {
                    let leaf=wanted.rsplit('\\').next().unwrap_or_default();
                    for path in self.child_folders(parent) {
                        let name=file_name(&path);
                        if matches_pattern(name,leaf) {entries.insert(name.to_lowercase(),(name.to_string(),self.attributes.get(&path).copied().unwrap_or(16)));}
                    }
                    if parent!=self.volume_root(&format!("{parent}\\")) {
                        for name in [".",".."] {
                            if matches_pattern(name,leaf) {
                                let key=if name=="." {parent.to_string()} else {self.key(&format!("{parent}\\.."))};
                                entries.insert(name.to_string(),(name.to_string(),self.attributes.get(&key).copied().unwrap_or(16)));
                            }
                        }
                    }
                }
            }
            let mut listing:Vec<_>=entries.into_values().filter(|(_,held)| {
                (held & 2==0 || attributes & 2!=0) && (held & 4==0 || attributes & 4!=0)
            }).map(|(name,_)|name).collect();
            listing.reverse();self.listing=Some(listing);
        }
        let Some(listing)=self.listing.as_mut() else {return Err(FileError::InvalidArgument);};
        if let Some(name)=listing.pop() {Ok(name)} else {self.listing=None;Ok(String::new())}
    }

    /// `Kill`: every file answering the pattern.
    pub fn kill(&mut self, pattern: &str) -> Result<(), FileError> {
        self.delete_matching_files(pattern,false,false)
    }

    pub fn delete_files(&mut self, pattern: &str, force: bool) -> Result<(), FileError> {
        self.delete_matching_files(pattern,force,true)
    }

    fn delete_matching_files(&mut self, pattern: &str, force: bool, fso: bool) -> Result<(), FileError> {
        let wanted = self.key(pattern);
        let folder = wanted.rfind('\\').map_or(String::new(), |at| wanted[..=at].to_string());
        let names = self.names_matching(pattern);
        if names.is_empty() {
            return Err(FileError::NotFound);
        }
        for name in names {
            let full = self.key(&format!("{folder}{name}"));
            if self.open.values().any(|handle| handle.path == full) {
                return Err(if fso {FileError::PermissionDenied} else {FileError::AlreadyOpen});
            }
            if !force && self.is_read_only_key(&full) {return Err(if fso {FileError::PermissionDenied} else {FileError::PathFileAccess});}
            self.contents.remove(&full);
            self.attributes.remove(&full);
        }
        Ok(())
    }

    /// `FileCopy`.
    pub fn copy(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        if source.is_empty() {return Err(FileError::PathFileAccess);}
        self.copy_with_options(source,destination,false)
    }

    pub fn copy_fso(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        self.copy_fso_with_overwrite(source,destination,true)
    }

    pub fn copy_fso_with_overwrite(&mut self, source: &str, destination: &str, overwrite: bool) -> Result<(), FileError> {
        if source.is_empty() || destination.is_empty() {return Err(FileError::InvalidArgument);}
        let sources=self.fso_source_files(source)?;
        let folder_target=destination.ends_with(['\\','/']) || source.contains(['*','?']) && self.folder_exists(destination);
        if folder_target && !self.folder_exists(destination) {return Err(FileError::PathNotFound);}
        for source in sources {
            let destination=if folder_target {format!("{}\\{}",destination.trim_end_matches(['\\','/']),file_name(&source))} else {destination.to_string()};
            self.copy_one(&source,&destination,true,overwrite)?;
        }
        Ok(())
    }

    fn copy_with_options(&mut self, source: &str, destination: &str, fso: bool) -> Result<(), FileError> {
        self.copy_one(source,destination,fso,true)
    }

    fn copy_one(&mut self, source: &str, destination: &str, fso: bool, overwrite: bool) -> Result<(), FileError> {
        let bytes = self.contents.get(&self.key(source)).ok_or(FileError::NotFound)?.1.clone();
        let key=self.key(destination);
        let source_key=self.key(source);
        if fso {
            if self.sharing_conflict(&source_key,true,false,true,true) || self.sharing_conflict(&key,false,true,true,true) {return Err(FileError::PermissionDenied);}
        } else if self.open.values().any(|held|held.path==source_key || held.path==key) {
            return Err(FileError::AlreadyOpen);
        }
        if fso && !overwrite && self.contents.contains_key(&key) {return Err(FileError::AlreadyExists);}
        if !fso && destination.ends_with(['\\','/']) {return Err(FileError::PathNotFound);}
        if source_key==key {return if fso {Ok(())}else{Err(FileError::PermissionDenied)};}
        if let Some((parent,_))=key.rsplit_once('\\') {
            if !self.folder_exists(&format!("{parent}\\")) {return Err(FileError::PathNotFound);}
        }
        if self.folder_exists(&format!("{key}\\")) {return Err(if fso {FileError::PermissionDenied}else{FileError::PathFileAccess});}
        if self.is_read_only_key(&key) {return Err(if fso {FileError::PermissionDenied}else{FileError::PathFileAccess});}
        let attributes=if fso {self.attributes(source)? | 32}else{32};
        self.attributes.insert(key.clone(),attributes);
        self.contents.insert(key, (self.absolute_path(destination.trim()), bytes));
        Ok(())
    }

    /// `Name ... As ...`.
    pub fn rename(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        if source.is_empty() {return Err(FileError::PathFileAccess);}
        self.rename_with_options(source,destination,false)
    }

    pub fn rename_fso(&mut self, source: &str, destination: &str) -> Result<(), FileError> {
        if source.is_empty() || destination.is_empty() {return Err(FileError::InvalidArgument);}
        let sources=self.fso_source_files(source)?;
        let folder_target=destination.ends_with(['\\','/']) || source.contains(['*','?']) && self.folder_exists(destination);
        if folder_target && !self.folder_exists(destination) {return Err(FileError::PathNotFound);}
        for source in sources {
            let destination=if folder_target {format!("{}\\{}",destination.trim_end_matches(['\\','/']),file_name(&source))}else{destination.to_string()};
            self.rename_with_options(&source,&destination,true)?;
        }
        Ok(())
    }

    fn fso_source_files(&self, source: &str) -> Result<Vec<String>,FileError> {
        if !source.contains(['*','?']) {
            let key=self.key(source);
            return self.contents.get(&key).map(|(path,_)|vec![path.clone()]).ok_or(FileError::NotFound);
        }
        let pattern=self.absolute_path(source);let parent=pattern.rsplit_once('\\').map_or("",|(parent,_)|parent);
        let names=self.names_matching(source);
        if names.is_empty() {return Err(FileError::NotFound);}
        Ok(names.into_iter().map(|name|format!("{parent}\\{name}")).collect())
    }

    fn rename_with_options(&mut self, source: &str, destination: &str, fso: bool) -> Result<(), FileError> {
        let key=self.key(destination);let old=self.key(source);
        if old!=key && self.contents.contains_key(&key) {
            return Err(FileError::AlreadyExists);
        }
        let attributes=self.attributes(source)? | 32;
        if self.open.values().any(|held|held.path==old) {return Err(if fso {FileError::PermissionDenied} else {FileError::AlreadyOpen});}
        if !self.contents.contains_key(&old) {return Err(FileError::NotFound);}
        if old==key {
            let destination=self.absolute_path(destination.trim());
            if let Some((path,_))=self.contents.get_mut(&old) {
                if *path!=destination {*path=destination;self.attributes.insert(old,attributes);}
            }
            return Ok(());
        }
        if let Some((parent,_))=key.rsplit_once('\\') {
            if !self.folder_exists(&format!("{parent}\\")) {return Err(if fso {FileError::PathNotFound}else{FileError::NotFound});}
        }
        if self.folder_exists(&format!("{key}\\")) {return Err(FileError::PermissionDenied);}
        let (_, bytes) = self.contents.remove(&old).ok_or(FileError::NotFound)?;
        self.attributes.remove(&old);
        self.attributes.insert(key.clone(),attributes);
        self.contents.insert(key, (self.absolute_path(destination.trim()), bytes));
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_file_paths_preserve_api_errors_and_existing_bytes() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        assert_eq!(files.copy("","C:\\work\\copy"),Err(FileError::PathFileAccess));
        assert_eq!(files.rename("","C:\\work\\moved"),Err(FileError::PathFileAccess));
        assert_eq!(files.copy_fso("","C:\\work\\copy"),Err(FileError::InvalidArgument));
        assert_eq!(files.copy_fso("C:\\work\\held",""),Err(FileError::InvalidArgument));
        assert_eq!(files.rename_fso("","C:\\work\\moved"),Err(FileError::InvalidArgument));
        assert_eq!(files.rename_fso("C:\\work\\held",""),Err(FileError::InvalidArgument));
        assert_eq!(files.bytes("C:\\work\\held").unwrap(),b"seed");assert!(!files.exists("C:\\work\\copy"));assert!(!files.exists("C:\\work\\moved"));
    }

    #[test]
    fn fso_copy_preserves_source_flags_expands_names_and_honors_overwrite() {
        let mut files=Files::default();files.mount_folder("C:\\work\\out").unwrap();
        files.replace_bytes("C:\\work\\UPPER.TXT",b"seed".to_vec()).unwrap();files.set_attributes("C:\\work\\UPPER.TXT",1).unwrap();
        files.copy_fso("C:\\work\\*.txt","C:\\work\\out\\").unwrap();
        assert_eq!(files.bytes("C:\\work\\out\\UPPER.TXT").unwrap(),b"seed");assert_eq!(files.attributes("C:\\work\\out\\UPPER.TXT"),Ok(33));
        files.set_attributes("C:\\work\\UPPER.TXT",32).unwrap();files.set_attributes("C:\\work\\out\\UPPER.TXT",32).unwrap();
        assert_eq!(files.copy_fso_with_overwrite("C:\\work\\UPPER.TXT","C:\\work\\out\\",false),Err(FileError::AlreadyExists));
        files.set_attributes("C:\\work\\out\\UPPER.TXT",1).unwrap();
        assert_eq!(files.copy_fso("C:\\work\\UPPER.TXT","C:\\work\\out\\"),Err(FileError::PermissionDenied));
        files.set_attributes("C:\\work\\UPPER.TXT",1).unwrap();
        files.copy("C:\\work\\UPPER.TXT","C:\\work\\basic").unwrap();assert_eq!(files.attributes("C:\\work\\basic"),Ok(32));
        assert_eq!(files.copy_fso("C:\\work\\UPPER.TXT","C:\\missing\\new"),Err(FileError::PathNotFound));
        assert!(!files.exists("C:\\missing\\new"));assert_eq!(files.bytes("C:\\work\\UPPER.TXT").unwrap(),b"seed");
    }

    #[test]
    fn move_self_is_noop_and_invalid_parents_do_not_consume_the_source() {
        let mut files=Files::default();files.mount_folder("C:\\work\\out").unwrap();files.replace_bytes("C:\\work\\held.txt",b"seed".to_vec()).unwrap();
        files.rename("C:\\work\\held.txt","C:\\work\\held.txt").unwrap();files.rename_fso("C:\\work\\held.txt","C:\\work\\held.txt").unwrap();
        assert_eq!(files.rename("C:\\work\\held.txt","C:\\work\\missing\\new"),Err(FileError::NotFound));
        assert_eq!(files.rename_fso("C:\\work\\held.txt","C:\\work\\missing\\new"),Err(FileError::PathNotFound));
        assert_eq!(files.bytes("C:\\work\\held.txt").unwrap(),b"seed");
        files.rename_fso("C:\\work\\*.txt","C:\\work\\out\\").unwrap();assert!(!files.exists("C:\\work\\held.txt"));assert_eq!(files.bytes("C:\\work\\out\\held.txt").unwrap(),b"seed");
    }

    #[test]
    fn open_file_copy_and_move_keep_basic_and_fso_errors_and_bytes() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();
        files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();files.open("C:\\work\\held",OpenMode::Binary,1).unwrap();
        assert_eq!(files.copy("C:\\work\\held","C:\\work\\copy"),Err(FileError::AlreadyOpen));
        assert!(!files.exists("C:\\work\\copy"));
        assert_eq!(files.rename("C:\\work\\held","C:\\work\\moved"),Err(FileError::AlreadyOpen));
        assert_eq!(files.rename_fso("C:\\work\\held","C:\\work\\moved"),Err(FileError::PermissionDenied));
        assert!(!files.exists("C:\\work\\moved"));assert_eq!(files.bytes("C:\\work\\held").unwrap(),b"seed");
        files.copy_fso("C:\\work\\held","C:\\work\\copy").unwrap();assert_eq!(files.bytes("C:\\work\\copy").unwrap(),b"seed");
        files.close(&[]);files.replace_bytes("C:\\work\\copy",b"old".to_vec()).unwrap();files.open("C:\\work\\copy",OpenMode::Binary,1).unwrap();
        assert_eq!(files.copy("C:\\work\\held","C:\\work\\copy"),Err(FileError::AlreadyOpen));
        assert_eq!(files.bytes("C:\\work\\copy").unwrap(),b"old");
        files.copy_fso("C:\\work\\held","C:\\work\\copy").unwrap();
        assert_eq!(files.bytes("C:\\work\\copy").unwrap(),b"seed");
    }

    #[test]
    fn fso_writers_share_readers_but_refuse_other_writers_before_truncation() {
        use crate::ast::{FileAccess,FileLock};
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        files.open_with_options("C:\\work\\held",OpenMode::Binary,1,Some(FileAccess::Write),Some(FileLock::Shared)).unwrap();
        assert_eq!(files.open_stream("C:\\work\\held",OpenMode::Output),Err(FileError::PermissionDenied));
        assert_eq!(files.bytes("C:\\work\\held").unwrap(),b"seed");
        let reader=files.open_stream("C:\\work\\held",OpenMode::Input).unwrap();files.close(&[1,reader]);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,1,Some(FileAccess::Read),Some(FileLock::Shared)).unwrap();
        let writer=files.open_stream("C:\\work\\held",OpenMode::Append).unwrap();
        assert_eq!(files.open("C:\\work\\held",OpenMode::Binary,2),Err(FileError::PermissionDenied));
        let reader=files.open_stream("C:\\work\\held",OpenMode::Input).unwrap();files.close(&[writer,reader]);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Write),Some(FileLock::Shared)).unwrap();
    }

    #[test]
    fn sequential_and_share_conflicts_are_refused_before_replacement() {
        use crate::ast::{FileAccess,FileLock};
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        files.open("C:\\work\\held",OpenMode::Input,1).unwrap();
        assert_eq!(files.open("C:\\work\\held",OpenMode::Output,2),Err(FileError::AlreadyOpen));
        assert_eq!(files.bytes("C:\\work\\held"),Ok(b"seed".as_slice()));files.close(&[]);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,1,Some(FileAccess::Read),Some(FileLock::Write)).unwrap();
        assert_eq!(files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Write),Some(FileLock::Shared)),Err(FileError::PermissionDenied));
        assert_eq!(files.free_number(false),2);
        files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Read),Some(FileLock::Shared)).unwrap();files.close(&[2]);
        assert_eq!(files.open_with_options("C:\\work\\held",OpenMode::Binary,2,Some(FileAccess::Read),Some(FileLock::ReadWrite)),Err(FileError::PermissionDenied));
    }

    #[test]
    fn record_lock_ownership_exact_unlock_and_close_release_are_preserved() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();files.replace_bytes("C:\\work\\held",b"seed".to_vec()).unwrap();
        files.open("C:\\work\\held",OpenMode::Binary,1).unwrap();files.open("C:\\work\\held",OpenMode::Binary,2).unwrap();
        files.lock_records(1,Some(1),Some(2),false).unwrap();
        assert_eq!(files.lock_records(1,Some(1),Some(2),false),Err(FileError::PermissionDenied));
        assert_eq!(files.lock_records(1,Some(1),Some(1),true),Err(FileError::PermissionDenied));
        assert_eq!(files.get(2,1),Err(FileError::PermissionDenied));assert_eq!(files.next_position(2),Ok(1));
        assert_eq!(files.put(2,b"X"),Err(FileError::PermissionDenied));assert_eq!(files.bytes("C:\\work\\held"),Ok(b"seed".as_slice()));
        assert_eq!(files.get(1,1),Ok(b"s".to_vec()));
        files.close(&[1]);assert_eq!(files.get(2,1),Ok(b"s".to_vec()));
        assert_eq!(files.lock_records(2,Some(0),None,false),Err(FileError::BadRecordNumber));
        assert_eq!(files.lock_records(2,Some(4),Some(2),false),Err(FileError::BadRecordNumber));
        files.lock_records(2,None,None,false).unwrap();let state=files.closed_snapshot();files.restore_closed_snapshot(&state).unwrap();
        files.open("C:\\work\\held",OpenMode::Binary,1).unwrap();files.lock_records(1,None,None,false).unwrap();
    }

    #[test]
    fn dir_cursor_distinguishes_terminal_empty_and_invalid_continuation() {
        let mut files=Files::default();files.mount_folder("C:\\work\\empty").unwrap();files.change_directory("C:\\work").unwrap();
        files.replace_bytes("held.txt",vec![]).unwrap();
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some("held.txt"),0),Ok("held.txt".into()));
        assert_eq!(files.dir_checked(None,0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some("missing*"),0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some(""),0),Ok("held.txt".into()));
        assert_eq!(files.dir_checked(Some("empty"),16),Ok("empty".into()));
        assert_eq!(files.dir_checked(Some("*"),16),Ok(".".into()));
        assert_eq!(files.dir_checked(None,0),Ok("..".into()));
        assert_eq!(files.dir_checked(None,0),Ok("empty".into()));
    }

    #[test]
    fn volume_queries_use_host_facts_and_have_their_own_terminal_empty() {
        let mut files=Files::default();
        files.restore_volume_facts(&[("C:\\".into(),String::new()),("D:\\".into(),"Fixture volume".into())]).unwrap();
        assert_eq!(files.dir_checked(Some("D:\\held.txt"),8),Ok("Fixture volume".into()));
        assert_eq!(files.dir_checked(None,0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        assert_eq!(files.dir_checked(Some("C:\\held.txt"),8),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Ok(String::new()));
        assert_eq!(files.dir_checked(None,0),Err(FileError::InvalidArgument));
        let state=files.volume_facts();
        for invalid in [vec![("relative".into(),"x".into())],vec![("C:\\folder".into(),"x".into())],vec![("C:\\".into(),"x".into()),("c:/".into(),"y".into())]] {
            assert_eq!(files.restore_volume_facts(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(files.volume_facts(),state);
        }
    }

    #[test]
    fn file_access_refusals_preserve_bytes_and_handle_positions() {
        use crate::ast::FileAccess;
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();
        files.replace_bytes("C:\\work\\held.txt",b"seed".to_vec()).unwrap();
        files.set_attributes("C:\\work\\held.txt",1).unwrap();
        for mode in [OpenMode::Binary,OpenMode::Random(4)] {
            files.open("C:\\work\\held.txt",mode,1).unwrap();
            assert_eq!(files.put(1,b"X"),Err(FileError::PathFileAccess));
            assert_eq!(files.next_position(1),Ok(1));
            assert_eq!(files.bytes("C:\\work\\held.txt"),Ok(b"seed".as_slice()));
            assert_eq!(files.get(1,2),Ok(b"se".to_vec()));files.close(&[1]);
        }
        for access in [FileAccess::Write,FileAccess::ReadWrite] {
            assert_eq!(files.open_with_access("C:\\work\\held.txt",OpenMode::Binary,1,Some(access)),Err(FileError::PathFileAccess));
            assert_eq!(files.free_number(false),1);
        }
        files.set_attributes("C:\\work\\held.txt",0).unwrap();
        files.open_with_access("C:\\work\\held.txt",OpenMode::Binary,1,Some(FileAccess::Read)).unwrap();
        assert_eq!(files.put(1,b"X"),Err(FileError::PathFileAccess));
        assert_eq!(files.bytes("C:\\work\\held.txt"),Ok(b"seed".as_slice()));
    }

    #[test]
    fn dir_flags_filter_hidden_and_system_without_losing_cursor_state() {
        let mut files=Files::default();files.mount_folder("C:\\work").unwrap();
        for (name,attributes) in [("a.txt",1),("b.txt",2),("c.txt",4),("d.txt",0)] {
            let path=format!("C:\\work\\{name}");files.replace_bytes(&path,vec![]).unwrap();files.set_attributes(&path,attributes).unwrap();
        }
        assert_eq!(files.dir(Some("C:\\work\\*.txt")),"a.txt");
        assert_eq!(files.dir(None),"d.txt");assert_eq!(files.dir(None),"");
        assert_eq!(files.dir_with_attributes(Some("C:\\work\\b.txt"),2),"b.txt");
        assert_eq!(files.dir_with_attributes(Some("C:\\work\\c.txt"),4),"c.txt");
        assert_eq!(files.dir_with_attributes(Some("C:\\work\\d.txt"),1),"d.txt");
    }

    #[test]
    fn attributes_survive_transport_and_refused_updates_are_atomic() {
        let mut files=Files::default();files.mount_folder("C:\\work\\empty").unwrap();
        files.replace_bytes("C:\\work\\held.txt",b"held".to_vec()).unwrap();
        files.set_attributes("C:\\work\\held.txt",1).unwrap();
        files.set_fso_attributes("C:\\work\\empty",-1).unwrap();
        let state=files.closed_snapshot();
        let mut restored=Files::default();restored.restore_closed_snapshot(&state).unwrap();
        assert_eq!(restored.attributes("c:/WORK/held.txt"),Ok(1));
        assert_eq!(restored.attributes("c:/WORK/empty"),Ok(55));
        assert_eq!(restored.replace_bytes("C:\\work\\held.txt",vec![]),Err(FileError::PermissionDenied));
        assert_eq!(restored.kill("C:\\work\\held.txt"),Err(FileError::PathFileAccess));
        assert_eq!(restored.closed_snapshot(),state);
        for attributes in [vec![("C:\\work\\missing".into(),0)],vec![("C:\\work\\held.txt".into(),16)],
            vec![("C:\\work\\empty".into(),0)],vec![("C:\\work\\held.txt".into(),64)],
            vec![("C:\\work\\held.txt".into(),-1)],vec![("C:\\work\\held.txt".into(),1),("c:/WORK/HELD.txt".into(),0)]] {
            let invalid=FileSystemSnapshot {attributes,..state.clone()};
            assert_eq!(restored.restore_closed_snapshot(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(restored.closed_snapshot(),state);
        }
        // Legacy snapshots without overrides replace metadata with defaults.
        restored.restore_closed_snapshot(&FileSystemSnapshot {attributes:vec![],..state.clone()}).unwrap();
        assert_eq!(restored.attributes("C:\\work\\held.txt"),Ok(32));
        assert_eq!(restored.attributes("C:\\work\\empty"),Ok(16));
    }

    #[test]
    fn readonly_directory_failure_retains_native_partial_deletion() {
        let mut files=Files::default();files.mount_folder("C:\\work\\tree\\child").unwrap();
        files.replace_bytes("C:\\work\\tree\\child\\nested.txt",vec![1]).unwrap();
        files.replace_bytes("C:\\work\\tree\\leaf.txt",vec![2]).unwrap();
        files.set_attributes("C:\\work\\tree\\child",1).unwrap();
        assert_eq!(files.delete_folder_with_force("C:\\work\\tree",false),Err(FileError::PermissionDenied));
        assert!(!files.exists("C:\\work\\tree\\child\\nested.txt"));
        assert!(files.exists("C:\\work\\tree\\leaf.txt"));
        assert_eq!(files.attributes("C:\\work\\tree\\child"),Ok(17));
        files.delete_folder_with_force("C:\\work\\tree",true).unwrap();
        assert!(!files.folder_exists("C:\\work\\tree"));
        assert!(!files.closed_snapshot().attributes.iter().any(|(path,_)|path.to_lowercase().contains("tree")));
    }

    #[test]
    fn directory_operations_preserve_refused_writes_and_native_partial_deletion_order() {
        let mut files=Files::default();
        for folder in ["C:\\work\\empty","C:\\work\\exists","C:\\work\\nonempty\\child","C:\\work\\nonempty\\empty-child"] {
            files.mount_folder(folder).unwrap();
        }
        for (path,bytes) in [("C:\\work\\held.txt",b"held".to_vec()),("C:\\work\\nonempty\\leaf.txt",b"leaf".to_vec()),("C:\\work\\nonempty\\child\\nested.txt",b"nested".to_vec())] {
            files.replace_bytes(path,bytes).unwrap();
        }
        let before=files.closed_snapshot();
        assert_eq!(files.make_folder("C:\\work\\missing\\child"),Err(FileError::PathNotFound));
        assert_eq!(files.make_folder("C:\\work\\held.txt"),Err(FileError::AlreadyExists));
        assert_eq!(files.remove_folder("C:\\work\\nonempty",false),Err(FileError::PathFileAccess));
        assert_eq!(files.closed_snapshot(),before);
        // Native case53: sorted enumeration reaches a locked ordinary file
        // after the earlier empty directories, before the later subtree.
        files.open("C:\\work\\held.txt",OpenMode::Input,1).unwrap();
        assert_eq!(files.remove_folder("C:\\work",true),Err(FileError::PermissionDenied));
        assert!(!files.folder_exists("C:\\work\\empty"));
        assert!(!files.folder_exists("C:\\work\\exists"));
        assert_eq!(files.bytes("C:\\work\\nonempty\\child\\nested.txt").unwrap(),b"nested");
        files.close(&[1]);
        // Native case44: earlier child directories are removed, while the
        // locked leaf and its containing directory remain available.
        files.open("C:\\work\\nonempty\\leaf.txt",OpenMode::Input,1).unwrap();
        assert_eq!(files.remove_folder("C:\\work\\nonempty",true),Err(FileError::PermissionDenied));
        assert!(!files.folder_exists("C:\\work\\nonempty\\child"));
        assert!(!files.folder_exists("C:\\work\\nonempty\\empty-child"));
        assert_eq!(files.bytes("C:\\work\\nonempty\\leaf.txt").unwrap(),b"leaf");
        assert!(files.folder_exists("C:\\work\\nonempty"));
    }

    #[test]
    fn closed_namespace_snapshots_preserve_empty_folders_and_bytes_atomically() {
        let mut files=Files::default();
        files.mount_folder("C:\\work\\empty").unwrap();
        files.change_directory("C:\\work").unwrap();
        files.replace_bytes("binary.dat",vec![0,255,128,1]).unwrap();
        files.open("binary.dat",OpenMode::Input,1).unwrap();
        let state=files.closed_snapshot();
        for invalid in [
            FileSystemSnapshot {folders:vec!["relative".into()],files:state.files.clone(),attributes:vec![]},
            FileSystemSnapshot {folders:vec!["C:\\work".into(),"c:/WORK".into()],files:vec![],attributes:vec![]},
            FileSystemSnapshot {folders:vec![],files:vec![("C:\\same".into(),vec![1]),("c:/SAME".into(),vec![2])],attributes:vec![]},
            FileSystemSnapshot {folders:vec![],files:vec![("C:\\bad|name".into(),vec![])],attributes:vec![]},
        ] {
            assert_eq!(files.restore_closed_snapshot(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(files.closed_snapshot(),state);
            assert!(files.handle(1).is_ok());
        }
        let mut restored=Files::default();restored.restore_closed_snapshot(&state).unwrap();
        assert!(restored.folder_exists("c:/WORK/EMPTY"));
        assert_eq!(restored.bytes("C:\\work\\binary.dat").unwrap(),&[0,255,128,1]);
        files.restore_closed_snapshot(&state).unwrap();
        assert_eq!(files.free_number(false),1);
        assert_eq!(files.current_directory(None).unwrap(),"C:\\work");
    }

    #[test]
    fn directory_snapshots_restore_drive_memory_and_reject_invalid_input_atomically() {
        let mut files=Files::default();
        files.mount_folder("C:\\first").unwrap();files.mount_folder("D:\\other").unwrap();
        files.change_directory("C:\\first").unwrap();files.change_directory("D:\\other").unwrap();files.change_drive("D").unwrap();
        let state=files.directory_state();
        let mut restored=Files::default();restored.restore_directory_state(&state).unwrap();
        assert_eq!(restored.current_directory(None).unwrap(),"D:\\other");
        assert_eq!(restored.absolute_path("C:held.txt"),"C:\\first\\held.txt");
        for invalid in [
            DirectoryState {current_drive:'Z',directories:state.directories.clone()},
            DirectoryState {current_drive:'?',directories:state.directories.clone()},
            DirectoryState {current_drive:'C',directories:vec!["relative".into()]},
            DirectoryState {current_drive:'C',directories:vec!["C:\\first".into(),"c:\\other".into()]},
            DirectoryState {current_drive:'C',directories:vec!["C:\\bad|path".into()]},
        ] {
            assert_eq!(restored.restore_directory_state(&invalid),Err(FileError::InvalidArgument));
            assert_eq!(restored.directory_state(),state);
        }
    }

    #[test]
    fn file_exports_and_folder_listings_keep_their_resolved_paths() {
        let mut files=Files::default();
        files.mount_folder("C:\\work\\sub").unwrap();
        let before=files.clone();
        files.change_directory("C:\\work").unwrap();
        files.open("held.txt",OpenMode::Output,1).unwrap();
        files.write(1,"ROOT").unwrap();files.close(&[1]);
        files.change_directory("sub").unwrap();
        files.open("held.txt",OpenMode::Output,1).unwrap();
        files.write(1,"CHILD").unwrap();files.close(&[1]);
        assert_eq!(files.files_in("C:\\work"),vec!["C:\\work\\held.txt"]);
        assert_eq!(files.changed_files(&before),vec![("C:\\work\\held.txt".into(),b"ROOT".to_vec()),("C:\\work\\sub\\held.txt".into(),b"CHILD".to_vec())]);
        let previous=files.clone();
        files.rename("held.txt","renamed.txt").unwrap();
        assert_eq!(files.deleted_files(&previous),vec!["C:\\work\\sub\\held.txt"]);
        assert_eq!(files.changed_files(&previous),vec![("C:\\work\\sub\\renamed.txt".into(),b"CHILD".to_vec())]);
        files.copy("renamed.txt","..\\copied.txt").unwrap();
        assert_eq!(files.files_in("C:\\work"),vec!["C:\\work\\copied.txt","C:\\work\\held.txt"]);
    }

    #[test]
    fn current_directories_resolve_files_without_rebinding_open_handles() {
        let mut files=Files::default();
        for folder in ["C:\\first", "C:\\second", "D:\\other"] { files.mount_folder(folder).unwrap(); }
        files.change_directory("C:\\first").unwrap();
        files.open("value.txt",OpenMode::Output,1).unwrap();
        files.write(1,"first").unwrap();
        files.change_directory("C:\\second").unwrap();
        files.open(".\\value.txt",OpenMode::Output,2).unwrap();
        files.write(2,"second").unwrap();
        files.write(1,"!").unwrap();
        assert_eq!(files.bytes("..\\first\\value.txt").unwrap(),b"first!");
        assert_eq!(files.bytes("C:value.txt").unwrap(),b"second");
        assert_eq!(files.bytes("\\second\\value.txt").unwrap(),b"second");
        files.change_directory("D:\\other").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"C:\\second");
        files.change_drive("D:").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"D:\\other");
        files.change_directory("C:\\first").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"D:\\other");
        files.change_drive("C:\\ignored").unwrap();
        assert_eq!(files.current_directory(None).unwrap(),"C:\\first");
        assert_eq!(files.clone().current_directory(Some("D")).unwrap(),"D:\\other");
        assert_eq!(files.change_directory("missing"),Err(FileError::PathNotFound));
        assert_eq!(files.current_directory(None).unwrap(),"C:\\first");
        assert_eq!(files.change_drive("?"),Err(FileError::InvalidArgument));
        assert_eq!(files.current_directory(Some("Z")),Err(FileError::DeviceUnavailable));
    }

    #[test]
    fn host_binary_output_is_immediately_readable_and_replacement_is_atomic() {
        let mut files = Files::default();
        assert_eq!(files.replace_bytes("C:\\missing\\copy.xlsm", vec![1]), Err(FileError::PathNotFound));
        assert!(!files.exists("C:\\missing\\copy.xlsm"));
        files.make_folder("C:\\copies").unwrap();
        files.replace_bytes("C:\\copies\\copy.xlsm", vec![b'P', b'K', 0, 255]).unwrap();
        assert_eq!(files.length("c:/COPIES/COPY.xlsm"), Ok(4));
        assert_eq!(files.dir(Some("C:\\copies\\*.xlsm")), "copy.xlsm");
        files.open("c:/copies/copy.xlsm", OpenMode::Binary, 1).unwrap();
        assert_eq!(files.get(1, 4).unwrap(), vec![b'P', b'K', 0, 255]);
        assert_eq!(files.replace_bytes("C:\\copies\\copy.xlsm", vec![2]), Err(FileError::AlreadyOpen));
        assert_eq!(files.bytes("C:\\copies\\copy.xlsm").unwrap(), &[b'P', b'K', 0, 255]);
        files.move_to(1, 3).unwrap();
        files.put(1, &[7]).unwrap();
        assert_eq!(files.bytes("C:\\copies\\copy.xlsm").unwrap(), &[b'P', b'K', 7, 255]);
        files.close(&[1]);
        files.replace_bytes("C:\\copies\\copy.xlsm", vec![2]).unwrap();
        assert_eq!(files.bytes("C:\\copies\\copy.xlsm").unwrap(), &[2]);
    }

    #[test]
    fn a_file_written_reads_back_line_by_line_and_field_by_field() {
        let mut files = Files::default();
        files.open("C:\\tmp\\a.txt", OpenMode::Output, 1).unwrap();
        files.write(1, "one\r\n\"q, x\",2.5,#TRUE#\r\n").unwrap();
        files.close(&[1]);
        assert_eq!(files.length("c:/TMP/A.TXT"), Ok(24));
        files.open("C:\\tmp\\a.txt", OpenMode::Input, 1).unwrap();
        assert_eq!(files.read_line(1).unwrap(), "one");
        assert_eq!(files.read_field(1).unwrap(), ("q, x".to_string(), true));
        assert_eq!(files.read_field(1).unwrap(), ("2.5".to_string(), false));
        assert_eq!(files.read_field(1).unwrap(), ("#TRUE#".to_string(), false));
        assert!(files.at_end(1).unwrap());
        assert_eq!(files.dir(Some("C:\\tmp\\*.txt")), "a.txt");
        assert_eq!(files.dir(None), "");
    }
}
