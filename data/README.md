# Data folder

One folder per client or test dataset, each with the same two subfolders:

```
data/
├── CLNT-1/
│   ├── Input/    <- drop this cycle's ONE input file here
│   └── Output/   <- the review writes every deliverable here
├── TEST-Meridian-FY2025/
│   ├── Input/
│   └── Output/
└── TEST-Halcyon-FY2025/
    ├── Input/
    └── Output/
```

- **New client:** create `<CLIENT>/Input` and `<CLIENT>/Output`, then set `"active_client": "<CLIENT>"` in
  `config/data-paths.json` (or just tell Claude which client to run).
- **Older inputs and runs** go in `_archive` subfolders. The pipeline ignores anything starting with `_`, so
  `Input/` still holds exactly one live file.
- Everything in this folder except this README is ignored by git: client data never gets committed.
