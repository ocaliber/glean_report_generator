# 📝 Report Generator Help

This help file is partially AI generated. Looking through the template reports is a great way of understanding how the report generator can be used. 

Hello! Welcome to the Caliber Design report generator help. Below you’ll find the fields and placeholders available in your Word template. Feel free to customize any of these – the generator will pick them up automatically.

---

## 1. Reporting Period

- **`{{ reporting_period }}`**  
  A human-readable date range (e.g., “2025-07-21 – 2025-07-27”).

---

## 2. Projects Overview

Your template receives a `projects` list; one entry per selected project.

| Field                       | Description                                          |
|-----------------------------|------------------------------------------------------|
| `code`                      | Project code or ID                                   |
| `name`                      | Project name                                         |
| `range_hours`               | Hours logged **during** this report period           |
| `total_hours`               | All-time hours on this project                       |
| `range_billable`            | Billable hours in the selected period                |
| `total_billable`            | All-time billable hours                              |
| `budget_used`               | Budget already spent (currency)                      |
| `budget_hours`              | Total budgeted hours                                 |
| `budget_amount`             | Budget amount (currency)                             |
| `monthly_user_budget_hours` | User’s monthly budgeted hours                        |
| `uninvoiced_amount`         | Dollars not yet invoiced                             |

**Example Jinja loop**:

| Code | Name | Period Hrs | Billable Hrs | Budget Used | Uninvoiced |
|------|------|------------|--------------|-------------|------------|
{% for p in projects %}
| {{ p.code }} | {{ p.name }} | {{ p.range_hours }} | {{ p.range_billable }} | ${{ p.budget_used }} | ${{ p.uninvoiced_amount }} |
{% endfor %}

---

## 3. Detailed Time Log

Use `entries`, a list of all time entries in the period.

| Field           | Description                                     |
|-----------------|-------------------------------------------------|
| `project`       | Project name                                    |
| `project_code`  | Project code                                    |
| `spent_date`    | Entry date (e.g., “2025-07-22”)                 |
| `task`          | Task name                                       |
| `notes`         | Free-text notes                                 |
| `rounded_hours` | Hours rounded to nearest quarter                |
| `billable`      | `true`/`false`                                  |
| `user.name`     | Person who logged the entry                     |

**Example**:

```jinja
{% for e in entries %}
| {{ e.spent_date }} | {{ e.project }} | {{ e.task }} | {{ e.notes }} | {{ e.rounded_hours }} hrs | {{ e.user.name }} |
{% endfor %}
```

---

## 4. People Summary

`people_summary` breaks down billable hours by user, then by project and task.

- **`name`**: User’s name  
- **`total`**: Total hours for this period  

Each person has a `projects` list:

| Field     | Description             |
|-----------|-------------------------|
| `project` | Project name            |
| `total`   | Hours on that project   |
| `tasks`   | List of `{ task, hours }` |

**Loop example**:

```jinja
{% for person in people_summary %}
### {{ person.name }} — {{ person.total }} hrs
{% for proj in person.projects %}
- **{{ proj.project }}** ({{ proj.total }} hrs)
  {% for t in proj.tasks %}
  - {{ t.task }}: {{ t.hours }} hrs
  {% endfor %}
{% endfor %}
{% endfor %}
```

---

## 5. AI Summarization (Optional)

To include an AI-generated summary, wrap your prompt between:

```text
[[AI_SUMMARY]]
Your custom instructions here…
[[END_AI_SUMMARY]]
```

- The generator grabs everything between those markers.  
- You control grouping, bullet style, length, etc.  
- Remove both tags if you don’t want an AI summary.

---

## Tips & Tricks

- **Currency formatting**: wrap values in `\$…` or use Word’s built-in formatting.    
- **Test early**: preview with sample data to catch any template typos.

> 💬 Thanks for using Caliber Design’s report generator!
