# Open a large workbook in the web application

`linexcel serve` copies a workbook into an index on disk at import. It then
calculates only the cells that you select. A large workbook therefore needs
disk space and time at import, and little memory. Rendering sheet captures is
the operation that needs memory.

Use this page when a workbook is larger than about 5 MB, or when an operation
stops on a limit.

## What an import costs

These imports ran on one Windows 11 machine with 16 logical processors, 31 GiB
of memory and an SSD. Cells with formulas cost time. Cells with plain values
cost disk space.

| Workbook | File size | Indexed cells | Import | Index on disk | Peak worker memory |
| --- | --- | --- | --- | --- | --- |
| Public transport model, `.xlsm` with VBA, 79,529 formulas | 6 MB | 125,365 | 1 min 10 s | 161 MiB | 93 MiB |
| Generated ledger and forecast, `.xlsm` with VBA | 20 MB | 2,709,922 | 4 min 51 s | 1.9 GiB | 89 MiB |
| The same model with more ledger rows, `.xlsm` with VBA | 50 MB | 6,773,384 | 11 min 9 s | 4.7 GiB | 117 MiB |

Plan for an index of 30 to 100 times the size of the file, and for 12 to
15 seconds of import for each megabyte. Larger files were not measured.

The 122 sheet captures of the 6 MB workbook took 2 min 43 s. LibreOffice
renders them, and it needed more than 1,024 MiB of memory.

## The limits that apply by default

When you start `linexcel serve` without limit options, the server reads the
machine and sets its own limits. It prints them under the address of the
application.

| Limit | Default | Option |
| --- | --- | --- |
| Memory for all running operations | Half of the physical memory in whole GiB, and at least 2,048 MiB | `--memory-mb` |
| Disk space for all workspaces | The space already used, plus the free space of that disk minus a reserve, in whole GiB. The reserve is 10% of the disk, and at most 10 GiB | `--max-storage-mb` |
| Size of one import, references included | 256 MiB | `--max-upload-mb` |
| Operations that run at the same time | 2 | `--workers` |

Each operation also has its own limits. The import form shows them above the
**Import** button, for example `Limits: 60 min · 120 s · 7.5 GB`. To change
them, select **Change** on that line, or **Settings** at the bottom of the
sidebar, then **Operation limits**.

| Field | Applies to | Default |
| --- | --- | --- |
| **Long operations** | Import, graph construction, sheet captures and investigation report | 60 minutes, which is also the maximum |
| **Targeted recalculation** | The calculation of one selected cell | 120 seconds |
| **Memory per operation** | Every operation | The server memory divided by the number of workers, in steps of 0.5 GB |

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
| `Time limit of … s reached` | The operation ran longer than its duration | Raise the field that the message names in **Settings**, then start the operation again |
| `Memory budget of … GB (… MiB) reached` | The operation needed more memory than its budget. Sheet captures of a large workbook are the usual case | Raise **Memory per operation**. To go above the server ceiling, restart the server with a larger `--memory-mb` |
| `Import too large` | The workbook and its references exceed the size of one import | Restart the server with a larger `--max-upload-mb` |

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
processors. A 50 MB workbook takes about 11 minutes on the machine above. You
can open another workspace and calculate its cells during that time. You can
also close the page. The import continues on the server, and the **Operations**
tab shows its state when you return.
