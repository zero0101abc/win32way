# main.py Flowchart

```
┌─────────────────────────────┐
│         Start               │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│  Load config (config.cfg)   │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│   Create Graph instance     │
└──────────────┬──────────────┘
               ▼
┌─────────────────────────────┐
│    greet_user(graph)        │
└──────────────┬──────────────┘
               ▼
        ┌─────────────┐
        │ choice = -1 │
        └──────┬──────┘
               ▼
    ┌─────────────────────┐
    │  while choice != 0  │──── No ────┐
    └──────────┬──────────┘             │
               │ Yes                    │
               ▼                        │
┌─────────────────────────────┐         │
│  Print menu options (0-4)   │         │
└──────────────┬──────────────┘         │
               ▼                         │
┌─────────────────────────────┐         │
│    Get user input           │         │
└──────────────┬──────────────┘         │
               ▼                         │
        ┌─────────────┐                 │
        │ Valid int?  │── No ──► choice=-1
        └──────┬──────┘                 │
               │ Yes                    │
               ▼                        │
    ┌─────────────────────┐             │
    │    choice == 0?     │── Yes ──► Goodbye... ──► Exit Loop
    └──────────┬──────────┘             │
               │ No                     │
               ▼                        │
    ┌─────────────────────┐             │
    │    choice == 1?     │── Yes ──► display_access_token()
    └──────────┬──────────┘             │
               │ No                     │
               ▼                        │
    ┌─────────────────────┐             │
    │    choice == 2?     │── Yes ──► list_inbox()
    └──────────┬──────────┘             │
               │ No                     │
               ▼                        │
    ┌─────────────────────┐             │
    │    choice == 3?     │── Yes ──► send_mail()
    └──────────┬──────────┘             │
               │ No                     │
               ▼                        │
    ┌─────────────────────┐             │
    │    choice == 4?     │── Yes ──► make_graph_call()
    └──────────┬──────────┘             │
               │ No                     │
               ▼                        │
    ┌─────────────────────┐             │
    │  Print "Invalid!"   │─────────────┘
    └─────────────────────┘
               │
               ▼ (if ODataError)
    ┌─────────────────────┐
    │  Print error details │
    └──────────┬──────────┘
               │
               ▼ (loop back)
        ┌─────────────┐
        │ choice = -1 │
        └──────┬──────┘
               │
               └──────────► [Back to while loop]
```

## Functions

| Function | Description |
|----------|-------------|
| `main()` | Entry point, loads config, runs CLI menu loop |
| `greet_user(graph)` | Currently returns immediately (TODO) |
| `display_access_token(graph)` | Calls `graph.get_user_token()` and prints it |
| `send_mail(graph)` | Currently returns immediately (TODO) |
| `make_graph_call(graph)` | Currently returns immediately (TODO) |
| `list_inbox(graph)` | Not defined in this file (in graph.py) |
