# Email Filter System Flowchart

```mermaid
flowchart TD
    A[Start] --> B[Initialize EmailFilterManager]
    B --> C[Load Filters from email_filters.json]
    C --> D[Connect to Outlook via COM]
    D --> E[Get Inbox Folder]
    E --> F[Iterate Through Emails]
    
    F --> G[Extract Email Data]
    G --> G1[Sender]
    G --> G2[Subject]
    G --> G3[Body]
    G --> G4[Date]
    G --> G5[Recipients]
    
    G1 --> H[Check for iSupport Recipient]
    G2 --> H
    G3 --> H
    G4 --> H
    G5 --> H
    
    H --> I{Has iSupport Recipient?}
    I -->|Yes| J[Keep Email]
    I -->|No| K[Apply Filters]
    
    K --> L[For Each Enabled Filter]
    L --> M{Check From Email Filter}
    M -->|Match| N{Check Subject Filter}
    M -->|No Match| O[Skip to Next Filter]
    
    N -->|Match| P{Check Body Filter}
    N -->|No Match| O
    
    P -->|Match| Q[Add Action to List]
    P -->|No Match| O
    
    O --> R{More Filters?}
    R -->|Yes| L
    R -->|No| S{Any Actions Matched?}
    
    Q --> R
    S -->|Yes| J
    S -->|No| T[Discard Email]
    
    J --> U[Add Filter Actions to Email Data]
    U --> V[Store in all_emails List]
    V --> W{More Emails?}
    W -->|Yes| F
    W -->|No| X[Save to outlook_emails.json]
    
    T --> W
    X --> Y[End]

    style A fill:#4CAF50,color:#fff
    style Y fill:#f44336,color:#fff
    style J fill:#2196F3,color:#fff
    style T fill:#FF9800,color:#fff
```

## Filter Processing Detail

```mermaid
flowchart TD
    A[Single Filter Processing] --> B[Filter Enabled?]
    B -->|No| C[Skip Filter]
    B -->|Yes| D[From Email Condition?]
    
    D -->|Empty| E[Subject Condition?]
    D -->|Not Empty| F{Sender Contains From Email?}
    F -->|No| C
    F -->|Yes| E
    
    E -->|Empty| G[Body Condition?]
    E -->|Not Empty| H[Evaluate Power Automate Expression]
    
    H --> I{Subject Expression True?}
    I -->|No| C
    I -->|Yes| G
    
    G -->|Empty| J[Add Action]
    G -->|Not Empty| K[Evaluate Power Automate Expression]
    
    K --> L{Body Expression True?}
    L -->|No| C
    L -->|Yes| J
    
    J --> M[Filter Matched!]
    
    style C fill:#FF9800,color:#fff
    style M fill:#4CAF50,color:#fff
```

## Power Automate Expression Evaluation

```mermaid
flowchart TD
    A[Expression: contains(subject, 'Alert')] --> B[Parse Function Call]
    B --> C[Extract Function: contains]
    B --> D[Extract Arguments: subject, 'Alert']
    
    D --> E[Replace Variables]
    E --> E1[subject → Actual Email Subject]
    E --> E2['Alert' → 'Alert']
    
    E1 --> F[Call contains function]
    F --> G{subject.lower().contains('alert'.lower())}
    G -->|True| H[Return True]
    G -->|False| I[Return False]
    
    style H fill:#4CAF50,color:#fff
    style I fill:#f44336,color:#fff
```

## Key Components

### 1. EmailFilterManager Class
- **Load Filters**: Reads `email_filters.json`
- **Apply Filters**: Tests each email against all enabled filters
- **Evaluate Expressions**: Processes Power Automate-like expressions

### 2. Filter Conditions
- **From Email**: Simple string match in sender field
- **Subject Filter**: Power Automate expression evaluation
- **Body Filter**: Power Automate expression evaluation

### 3. Power Automate Functions Supported
- `equals(value1, value2)`
- `contains(text, search)`
- `startsWith(text, prefix)`
- `endsWith(text, suffix)`
- `concat(*args)`
- `substring(text, start, length)`
- `indexOf(text, search)`
- `split(text, separator)`

### 4. Data Flow
1. **Input**: Outlook emails via COM
2. **Processing**: Filter application with expression evaluation
3. **Output**: JSON file with filtered emails and actions

### 5. Filter Storage
```json
{
    "id": 1,
    "name": "MX Alert Filter",
    "from_email": "system.MX@hkt-emsconnect.com",
    "subject_filter": "contains(subject, 'Alert')",
    "body_filter": "contains(body, 'error')",
    "action": "send_to_support_team",
    "enabled": true
}
```

## Decision Points

1. **iSupport Check**: Emails sent to "iSupport" are always kept
2. **Filter Matching**: Email must pass ALL conditions in a filter
3. **Multiple Filters**: Email can match multiple filters (AND logic)
4. **Action Collection**: All matching filter actions are collected