# Open a large workbook in the web application

`linexcel serve` copies a workbook into an index on disk at import. It then
calculates only the cells that you select. A large workbook therefore needs
disk space and time at import, and little memory. Rendering sheet captures is
the operation that needs memory.

Use this page when a workbook is larger than about 5 MB, or when an operation
stops on a limit.

## What an import costs

These imports ran on one Windows 11 machine with 16 logical processors, 31 GiB
of memory and an SSD. Cells with formulas cost time and disk space. Plain
values go in one grid per sheet and cost little.

| Workbook | File size | Indexed cells | Index on disk | Peak memory of the import |
| --- | --- | --- | --- | --- |
| Public transport model, `.xlsm` with VBA, 79,529 formulas | 6 MB | 125,365 | 134 MiB | 95 MiB |
| Generated ledger and forecast, `.xlsm` with VBA | 20 MB | 2,709,922 | 317 MiB | 101 MiB |
| The same model with more ledger rows, `.xlsm` with VBA | 50 MB | 6,773,384 | 793 MiB | 159 MiB |
| The same model, larger again, `.xlsm` with VBA | 89 MB | 11,970,736 | 1.1 GiB | 214 MiB |

Plan for an index of about 15 to 25 times the size of the file when the
workbook holds many formulas. Larger files were not measured.

Version 1.20.0 imported these four workbooks in 1 min 10 s, 4 min 51 s,
11 min 9 s and 21 min 6 s on that machine while it did nothing else. The
current code was not timed on a quiet machine. Compared with 1.20.0 under the
same load, it was about as fast on the 6 MB model and faster on the 20 MB
one. Plan for about 10 to 15 seconds of import for each megabyte.

The 122 sheet captures of the 6 MB workbook took 2 min 43 s. LibreOffice
renders them, and it needed more than 1,024 MiB of memory.

## The limits that apply by default

When you start `linexcel serve` without limit options, the server reads the
machine and sets its own limits. It prints them under the address of the
application:

```text
Limits: memory 15360 MiB for 2 concurrent operations, 7680 MiB each (--memory-mb); storage 119792 MiB, 0 MiB used (--max-storage-mb); 3600 s per import, graph, capture or report (--import-seconds), 120 s per targeted recalculation.
```

| Limit | Default | Option |
| --- | --- | --- |
| Memory for all running operations | Half of the physical memory in whole GiB, and at least 2,048 MiB | `--memory-mb` |
| Disk space for all workspaces | The space already used, plus the free space of that disk minus a reserve. The reserve is 10% of the disk, and at most 10 GiB | `--max-storage-mb` |
| Size of one import, references included | 256 MiB | `--max-upload-mb` |
| Duration of an import, a graph construction, sheet captures or a report | 3,600 seconds | `--import-seconds` |
| Operations that run at the same time | 2 | `--workers` |

Each operation also has its own limits. The import form shows them above the
**Import** button, for example `Limits: 60 min · 120 s · 7.5 GB`. To change
them, select **Change** on that line, or **Settings** at the bottom of the
sidebar, then **Operation limits**.

| Field | Applies to | Default |
| --- | --- | --- |
| **Long operations** | Import, graph construction, sheet captures and investigation report | The value of `--import-seconds`: 60 minutes unless the server sets another one. 60 minutes is also the maximum. The field accepts a duration from one second |
| **Targeted recalculation** | The calculation of one selected cell | 120 seconds |
| **Memory per operation** | Every operation | The server memory divided by the number of workers, in steps of 0.5 GB, and at least 1 GB |

The fields return to these values when you reload the page. **Restore
recommended values** does the same without a reload.

An application that you embed with `create_app()` does not read the machine.
It keeps the values that the host passes, and 2,048 MiB for memory and for
disk space when the host passes none.

## When a limit stops an operation

The **Operations** tab shows the limit that stopped the operation, the value of
that limit and what to change. The workbook and the earlier results stay
available.

| Message starts with | Cause | What to do |
| --- | --- | --- |
| `Server storage quota reached` | The workbooks, indexes and results of all workspaces fill the disk space that the server may use | Delete workspaces that you no longer need, or restart the server with a larger `--max-storage-mb`. Then start the operation again |
| `Time limit of … s reached` | The operation ran longer than its duration | Raise the field that the message names in **Settings**, then start the operation again. At 3,600 seconds the message says that this is the longest duration allowed, and no setting raises it |
| `Memory budget of … GB (… MiB) reached` | The operation needed more memory than its budget. Sheet captures of a large workbook are the usual case | Raise **Memory per operation**. To go above the server ceiling, restart the server with a larger `--memory-mb` |
| `Import too large` | The workbook and its references exceed the size of one import | Restart the server with a larger `--max-upload-mb` |
| `Import stopped: the … s budget ran out` | A step of the import that cannot be interrupted outlasted the duration. No new index was published, and the message says whether the previous one remains | Select **Reimport with a budget of … s** in **Diagnostics**, or raise **Long operations**. At 3,600 seconds no longer duration exists |

## When an import runs out of time

An import fits its duration. When time runs short it stops reading sheets,
then stops analyzing formulas, and publishes what it has. **Diagnostics** then
shows **Partial import**, the sheets that were not imported and the number of
formulas that were not analyzed. A calculation that needs a missing cell is
refused with its reason. It is never calculated on absent values.

**Diagnostics** also estimates the duration that the whole import needs, and
offers **Reimport with a budget of … s**. That button puts the duration in
**Long operations** and starts the import again. The estimate comes from what
this import read. When it read little, Diagnostics marks the estimate as
uncertain, and it can be too short: the reimport is then partial again and
proposes a longer duration.

In two tests on a busy machine, the 6 MB workbook above was given 12 seconds.
Both imports were published as partial with an uncertain estimate. One
proposed 180 seconds, and its reimport was complete. The other proposed
120 seconds, its reimport was partial again and proposed 240 seconds.

For the rules of the estimate, see [The disk index](cli.md#the-disk-index).

## Give the server more room

Pass the options when you start the server. This command allows 40 GiB of
disk space and 12 GiB of memory:

```Shell
uvx linexcel serve --open --max-storage-mb 40960 --memory-mb 12288
```

## Import from the disk of the server

For a file of tens of megabytes, skip the transfer through the browser. In
**Import a workbook**, set **Workbook source** to **File on the server** and
enter the absolute path of the file. The server copies the file from its own
disk. This choice is available when the server runs on your machine with the
standalone launcher.

## Free disk space

**Local storage**, under the list of workspaces, shows the disk space that the
workspaces use and the space that the server may use.

Workspaces stay on disk until you delete them. To delete a workspace, open it,
open **Actions**, then select **Delete**. To delete old workspaces at each
start, pass `--retention-days` with a number of days.

## What stays slow

The import reads every cell and every formula once. It does not use several
processors. 60 minutes is the longest duration of an operation: an import
that needs more publishes what it could import in 60 minutes, as a partial
import. If the rates above hold for larger files, that is a workbook of about
240 MB on that machine, and a smaller one on a slower machine. No file of that
size was measured.

An import of tens of megabytes takes several minutes. You can open another
workspace and calculate its cells during that time. You can also close the
page. The import continues on the server, and the **Operations** tab shows
its state when you return.

## After an upgrade

This version stores an index in a new format. A workspace imported by version
1.20.0 or earlier asks to be reimported. Its workbook is kept, and its old
index is deleted when the reimport starts.
