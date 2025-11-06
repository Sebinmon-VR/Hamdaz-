

---

# User Priority Scoring System

This module provides a system to calculate **priority scores** and **ranks** for users based on:

1. **Active tasks** (low load → higher priority).
2. **Idle time** (among users with similar load, users who have been idle longer get higher priority).

It is designed for assigning tasks in a fair and intelligent way, balancing workload and fairness.

---

## Functions

### 1. `calculate_priority_score(user_analytics)`

**Purpose:**
Compute a priority score for each user based on active tasks and last assignment date.

**Parameters:**

| Parameter                                           | Type           | Description                                                   |
| --------------------------------------------------- | -------------- | ------------------------------------------------------------- |
| `user_analytics`                                    | `pd.DataFrame` | A DataFrame containing per-user analytics. Must have columns: |
| `OngoingTasksCount` (number of active tasks)        |                |                                                               |
| `LastAssignedDate` (datetime of last task assigned) |                |                                                               |

**Returns:**
`pd.DataFrame` — copy of the input with an additional column:

* `PriorityScore` → tuple `(-active_tasks, days_since_last)` representing each user's priority.

**Logic:**

1. **Primary factor:** Active tasks

   * Fewer tasks → higher priority.
   * Represented as negative value for sorting: `-active_tasks`.
2. **Secondary factor:** Idle time

   * Users who haven’t received tasks recently → higher priority.
   * Represented as `days_since_last` (number of days since last assignment).
3. The tuple `(-active_tasks, days_since_last)` ensures **lexicographic priority**: low tasks first, then idle users.

**Example:**

```python
import pandas as pd
from datetime import datetime, timezone

df = pd.DataFrame([
    {"User": "Alice", "OngoingTasksCount": 0, "LastAssignedDate": "2025-11-05T08:00:00+00:00"},
    {"User": "Bob", "OngoingTasksCount": 2, "LastAssignedDate": "2025-11-01T09:00:00+00:00"}
])

df = calculate_priority_score(df)
```

Resulting `PriorityScore` column:

| User  | PriorityScore |
| ----- | ------------- |
| Alice | (0, 1.0)      |
| Bob   | (-2, 5.0)     |

---

### 2. `assign_priority_rank(user_analytics)`

**Purpose:**
Assign a **priority rank** to each user based on the computed priority score.

**Parameters:**

| Parameter        | Type           | Description                                        |
| ---------------- | -------------- | -------------------------------------------------- |
| `user_analytics` | `pd.DataFrame` | DataFrame returned by `calculate_priority_score()` |

**Returns:**
`pd.DataFrame` — copy of input with additional column:

* `PriorityRank` → integer rank (1 = highest priority, N = lowest priority).

**Logic:**

1. Sort users **descending** by `PriorityScore` (tuple sorting).
2. Assign rank: highest priority user → `1`, next → `2`, etc.

**Example:**

```python
df = assign_priority_rank(df)
```

Resulting DataFrame:

| User  | PriorityScore | PriorityRank |
| ----- | ------------- | ------------ |
| Alice | (0, 1.0)      | 1            |
| Bob   | (-2, 5.0)     | 2            |

---

## Key Points

* **Fair task assignment:**
  Users with fewer tasks are prioritized first, ensuring no one is overloaded.

* **Idle consideration:**
  Among users with similar active tasks, those who haven’t received tasks recently get assigned next.

* **Tuple-based scoring:**
  Lexicographic tuple `( -active_tasks, days_since_last )` ensures correct ranking without manual weighting.

* **Extensible:**
  You can easily adjust priority rules, e.g., assign weights to active tasks vs. idle time, or add other factors like skill match.

---

## Dependencies

```python
import pandas as pd
from datetime import datetime, timezone
```

---
