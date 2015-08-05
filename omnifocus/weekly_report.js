#!/usr/bin/env osascript -l JavaScript

start = new Date();
lastMon = start.getDate() - start.getDay() - 6;
start.setDate(lastMon);
start.setHours(0);
start.setMinutes(0);
start.setSeconds(0);
end = new Date(start);
end.setDate(start.getDate() + 7);
end.setSeconds(-1);
week = start.getFullYear() + '_' + (start.getMonth() + 1) + '_' + start.getDate();

ofs = Application('OmniFocus');

md = 'Title: Task Report\n';
md += 'Subtitle: ' + start.toString() + ' - ' + end.toString() + '\n';
md += '\n';
md += '# Weekly Task Report\n\n';
md += 'from *' + start.toDateString() + '* to *' + end.toDateString() + '*\n\n';
ofs.defaultDocument.flattenedProjects().forEach(function(p) {
  if (p.context() === null ||
      p.context().name() != 'Work' ||
      p.status() != 'active') {return;}
  md += '## ' + p.name() + '\n\n';
  p.flattenedTasks().forEach(function(t) {
    checkBox = '- [ ] ';
    if (t.completed()) {
      if (t.completionDate() < start || t.completionDate() > end) { return; }
      checkBox = '- [x] ';
    }

    md += checkBox + t.name() + '\n';
  });
  md += '\n';
});
// ObjC.import('Cocoa')
// str = $.NSString.alloc.initWithUTF8String(md)
// str.writeToFileAtomically('/Users/xiaoxing/Projects/osxauto/omnifocus/report.md', true)

app = Application.currentApplication();
app.includeStandardAdditions = true;
reportFile = app.chooseFileName({
  withPrompt: 'Pick where to save the report.',
  defaultName: week + '_report.md'
});
report = app.openForAccess(reportFile, {writePermission: true});
app.write(md, {to: report});
app.closeAccess(reportFile);
