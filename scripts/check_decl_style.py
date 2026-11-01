import re, os

files = [
    'BASICpy/vb/Builtin/MaxMinImpl.bas',
    'BASICpy/vb/Builtin/SumImpl.bas',
    'BASICpy/vb/Builtin/FloatImpl.bas',
    'BASICpy/vb/Builtin/ChrOrdImpl.bas',
    'BASICpy/vb/Builtin/ReprImpl.bas',
    'LegacyApp.VB6/Arrays.cls',
    'LegacyApp.VB6/BitAccessor.cls',
    'LegacyApp.VB6/CStopwatch.cls',
    'LegacyApp.VB6/CStopwatches.cls',
    'LegacyApp.VB6/CTimeSpan.cls',
    'LegacyApp.VB6/CTimeSpanFactory.cls',
    'LegacyApp.VB6/CTimeSpans.cls',
    'LegacyApp.VB6/FileTimeDateTime.cls',
    'LegacyApp.VB6/FileTimeDateTimeFactory.cls',
    'LegacyApp.VB6/MinMax.cls',
    'LegacyApp.VB6/Numerics.cls',
    'LegacyApp.VB6/OleDates.cls',
]

base = 'vb_classic/GuiApp/'
for f in files:
    path = os.path.join(base, f)
    if not os.path.exists(path):
        path = f
    if not os.path.exists(path):
        continue
    with open(path, 'r', errors='replace') as fh:
        lines = fh.readlines()

    for i, line in enumerate(lines):
        m = re.match(r'\s*Dim\s+(\w+)\s+As\s+\w+:\s*\1\s*=', line)
        if m:
            varname = m.group(1)
            used_soon = False
            for j in range(i + 1, min(i + 6, len(lines))):
                next_line = lines[j]
                stripped = next_line.strip()
                if stripped.startswith("'") or stripped == '':
                    continue
                if re.search(r'\b' + varname + r'\b', stripped):
                    used_soon = True
                    break
            if not used_soon:
                print(f'{f}:{i+1}: {line.rstrip()}  <- var {varname} not used within 5 lines')
