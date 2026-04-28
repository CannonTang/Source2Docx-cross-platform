using System;
using System.Collections.Generic;
using System.Linq;

namespace Source2Docx.Models;

internal sealed class CodeTypeOption
{
    public CodeTypeOption(string displayName, IEnumerable<string> extensions)
    {
        DisplayName = displayName;
        Extensions = new HashSet<string>(
            extensions.Select(static item => item.ToLowerInvariant()),
            StringComparer.OrdinalIgnoreCase);
    }

    public string DisplayName { get; }

    public HashSet<string> Extensions { get; }

    public string PickerPattern => string.Join(";", Extensions.Select(static extension => "*" + extension));

    public override string ToString()
    {
        return DisplayName;
    }
}
