# CNC Manufacturing Examples: Tech Holdings Berhad

Real-world CNC machining operations data for **Tech Holdings Berhad**, a Malaysian SME manufacturer operating three plants in Penang with 508 total employees.

**Business Context:**
- CNC machining division with 52 employees across 3 plants
- Equipment: 18 CNC machines (3-axis VMC, 4-axis HMC, 5-axis, lathes)
- Industries served: Aerospace, Medical Devices, Electronics, Oil & Gas, Automotive, Industrial
- Fiscal Year: FY2025 (Oct 2024 - Sep 2025, all amounts in RM)
- Critical Challenge: High first-article failure rates, manual validation processes, margin erosion

## Business Trajectory

The company faces operational inefficiencies driving financial decline, with attempted interventions in Q3-Q4:

| Quarter | Revenue | Jobs | Failure Rate | Scrap Rate | OEE | Status |
|---------|---------|------|--------------|------------|-----|--------|
| **Q1** (Oct-Dec 2024) | RM4.11M | 195 | 18% | 12% | 58% | Baseline crisis |
| **Q2** (Jan-Mar 2025) | RM3.45M | 162 | 21% | 15% | 53% | Worsening, customer losses |
| **Q3** (Apr-Jun 2025) | RM3.20M | 148 | 16% | 11% | 56% | Process improvements |
| **Q4** (Jul-Sep 2025) | RM3.35M | 155 | 15% | 10% | 58% | Fragile stabilization |

**4-Quarter Total: RM14.11M revenue, 660 jobs, -16% decline Q1→Q2 with partial Q4 recovery**

## Files

| File | Size | Rows | Use Case |
|------|------|------|----------|
| `job_orders.xlsx` | 75KB | 660 | Production orders master, customer/part details |
| `job_execution.xlsx` | 69KB | 806 | Execution runs with prove-out attempts |
| `program_validation.xlsx` | 53KB | 660 | CAM/G-code validation, error tracking |
| `quality_inspections.xlsx` | 222KB | 3,607 | Dimensional inspection measurements |
| `cost_analysis.xlsx` | 79KB | 660 | Job costing with estimate vs actual variance |
| `scrap_rework.xlsx` | 9KB | 64 | Scrap events with root cause analysis |
| `material_inventory.xlsx` | 41KB | 624 | Weekly raw material movements |
| `tooling_management.xlsx` | 6KB | 19 | Tool usage and costs by quarter |
| `machine_downtime.xlsx` | 21KB | 330 | Equipment failures and maintenance |
| `production_schedule.xlsx` | 51KB | 936 | Weekly planning vs actual adherence |
| `machines.xlsx` | 6KB | 18 | Equipment master data |
| `labor_tracking.xlsx` | 7KB | 52 | Workforce by role/skill/shift |
| `customer_orders.xlsx` | 8KB | 54 | Customer master with churn tracking |

**Total: ~7,500 operational records across 4 quarters**

## Generate Examples

```bash
cd examples/cnc
source ../../.venv/bin/activate
python create_cnc_examples.py
```

## Usage

```bash
uvx --from mcp-server-excel-sql mcp-excel --path examples/cnc --watch
```

## Analysis Sequences

Interactive journeys demonstrating how to uncover Tech Holdings' operational issues.

### Sequence 1: Prove-Out Failure Analysis → RM500-5000 Per Failure

**Objective:** Understand why first-article failure rates are destroying profitability

```
Step 1: "How many programs required multiple prove-out attempts?"
→ SELECT COUNT(*) as programs_with_failures
  FROM "cnc.program_validation.Programs"
  WHERE prove_out_attempts > 1;

Step 2: "What's the distribution of prove-out attempts?"
→ SELECT
    prove_out_attempts,
    COUNT(*) as program_count,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM "cnc.program_validation.Programs"), 1) as pct
  FROM "cnc.program_validation.Programs"
  GROUP BY prove_out_attempts
  ORDER BY prove_out_attempts;

Step 3: "Which programmer has the highest failure rate?"
→ SELECT
    programmer_name,
    programmer_skill,
    COUNT(*) as total_programs,
    SUM(CASE WHEN prove_out_attempts > 1 THEN 1 ELSE 0 END) as failures,
    ROUND(SUM(CASE WHEN prove_out_attempts > 1 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as failure_pct
  FROM "cnc.program_validation.Programs"
  GROUP BY programmer_name, programmer_skill
  ORDER BY failure_pct DESC;

Step 4: "What errors cause prove-out failures?"
→ SELECT
    errors_found,
    COUNT(*) as occurrence_count
  FROM "cnc.program_validation.Programs"
  WHERE errors_found IS NOT NULL
  GROUP BY errors_found
  ORDER BY occurrence_count DESC
  LIMIT 10;

Step 5: "Did simulation reduce failures in Q3-Q4?"
→ SELECT
    CASE WHEN simulation_performed THEN 'Simulated' ELSE 'No Simulation' END as validation,
    COUNT(*) as programs,
    AVG(prove_out_attempts) as avg_attempts,
    SUM(CASE WHEN prove_out_attempts > 1 THEN 1 ELSE 0 END) as failures
  FROM "cnc.program_validation.Programs"
  GROUP BY simulation_performed;

Step 6: "Total cost wasted on prove-out failures"
→ SELECT
    SUM((prove_out_attempts - 1) * cost_per_attempt_rm) as total_wasted_rm
  FROM "cnc.program_validation.Programs"
  WHERE prove_out_attempts > 1;
```

**Key Insight:** Programs without simulation have 2.5x higher failure rates, costing RM300K+ in Q1-Q2.

---

### Sequence 2: Scrap Analysis → 60-70% Preventable

**Objective:** Quantify preventable waste and identify root causes

```
Step 1: "How much material was scrapped across all quarters?"
→ SELECT
    SUM(quantity_scrapped) as total_scrapped_parts,
    SUM(material_cost_lost_rm) as total_material_loss_rm,
    SUM(total_cost_impact_rm) as total_cost_impact_rm
  FROM "cnc.scrap_rework.Scrap";

Step 2: "What are the top scrap reasons?"
→ SELECT
    scrap_reason,
    COUNT(*) as events,
    SUM(quantity_scrapped) as parts_scrapped,
    SUM(total_cost_impact_rm) as cost_impact_rm,
    SUM(CASE WHEN preventable THEN 1 ELSE 0 END) as preventable_events
  FROM "cnc.scrap_rework.Scrap"
  GROUP BY scrap_reason
  ORDER BY cost_impact_rm DESC;

Step 3: "How much scrap was preventable?"
→ SELECT
    preventable,
    COUNT(*) as events,
    SUM(total_cost_impact_rm) as cost_rm,
    ROUND(SUM(total_cost_impact_rm) * 100.0 / (SELECT SUM(total_cost_impact_rm) FROM "cnc.scrap_rework.Scrap"), 1) as pct_of_total
  FROM "cnc.scrap_rework.Scrap"
  GROUP BY preventable;

Step 4: "Scrap trend by quarter"
→ SELECT
    j.quarter,
    COUNT(*) as scrap_events,
    SUM(s.total_cost_impact_rm) as cost_impact_rm
  FROM "cnc.scrap_rework.Scrap" s
  JOIN "cnc.job_orders.Orders" j ON s.job_id = j.job_id
  GROUP BY j.quarter
  ORDER BY j.quarter;

Step 5: "Which jobs had the worst scrap?"
→ SELECT
    s.job_id,
    j.customer_name,
    j.part_type,
    j.complexity,
    s.scrap_reason,
    s.total_cost_impact_rm
  FROM "cnc.scrap_rework.Scrap" s
  JOIN "cnc.job_orders.Orders" j ON s.job_id = j.job_id
  ORDER BY s.total_cost_impact_rm DESC
  LIMIT 10;
```

**Key Insight:** 65% of scrap (RM250K+) is preventable - primarily dimensional errors and operator mistakes.

---

### Sequence 3: Cost Variance Analysis → 20-30% Over Budget

**Objective:** Understand why jobs cost more than estimated

```
Step 1: "How many jobs exceeded cost estimates?"
→ SELECT
    COUNT(*) as total_jobs,
    SUM(CASE WHEN cost_variance_rm > 0 THEN 1 ELSE 0 END) as over_budget_jobs,
    ROUND(SUM(CASE WHEN cost_variance_rm > 0 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as over_budget_pct
  FROM "cnc.cost_analysis.Job_Costing";

Step 2: "Average cost overrun by complexity"
→ SELECT
    j.complexity,
    COUNT(*) as jobs,
    AVG(c.cost_variance_pct) as avg_variance_pct,
    SUM(c.cost_variance_rm) as total_variance_rm
  FROM "cnc.cost_analysis.Job_Costing" c
  JOIN "cnc.job_orders.Orders" j ON c.job_id = j.job_id
  GROUP BY j.complexity
  ORDER BY avg_variance_pct DESC;

Step 3: "Which cost category has the most variance?"
→ SELECT
    'Material' as category,
    SUM(actual_material_rm - estimated_material_rm) as total_variance_rm,
    AVG((actual_material_rm - estimated_material_rm) / estimated_material_rm * 100) as avg_variance_pct
  FROM "cnc.cost_analysis.Job_Costing"
  WHERE estimated_material_rm > 0
UNION ALL
  SELECT
    'Labor',
    SUM(actual_labor_rm - estimated_labor_rm),
    AVG((actual_labor_rm - estimated_labor_rm) / estimated_labor_rm * 100)
  FROM "cnc.cost_analysis.Job_Costing"
  WHERE estimated_labor_rm > 0
UNION ALL
  SELECT
    'Machine',
    SUM(actual_machine_rm - estimated_machine_rm),
    AVG((actual_machine_rm - estimated_machine_rm) / estimated_machine_rm * 100)
  FROM "cnc.cost_analysis.Job_Costing"
  WHERE estimated_machine_rm > 0;

Step 4: "Profitability by customer"
→ SELECT
    j.customer_name,
    COUNT(*) as jobs,
    SUM(c.revenue_rm) as total_revenue_rm,
    SUM(c.gross_margin_rm) as total_margin_rm,
    AVG(c.gross_margin_pct) as avg_margin_pct
  FROM "cnc.cost_analysis.Job_Costing" c
  JOIN "cnc.job_orders.Orders" j ON c.job_id = j.job_id
  WHERE c.revenue_rm IS NOT NULL
  GROUP BY j.customer_name
  ORDER BY avg_margin_pct DESC
  LIMIT 10;

Step 5: "Jobs with negative margins"
→ SELECT
    c.job_id,
    j.customer_name,
    j.part_type,
    j.complexity_score,
    c.revenue_rm,
    c.actual_total_cost_rm,
    c.gross_margin_rm,
    c.gross_margin_pct
  FROM "cnc.cost_analysis.Job_Costing" c
  JOIN "cnc.job_orders.Orders" j ON c.job_id = j.job_id
  WHERE c.gross_margin_rm < 0
  ORDER BY c.gross_margin_rm ASC;
```

**Key Insight:** Complex jobs (score 7+) average 35% cost overrun, driven by labor variance from rework.

---

### Sequence 4: Machine Utilization → OEE 53-58%

**Objective:** Understand capacity utilization and downtime patterns

```
Step 1: "Downtime by type and quarter"
→ SELECT
    quarter,
    downtime_type,
    COUNT(*) as events,
    ROUND(SUM(duration_hours), 1) as total_hours,
    ROUND(AVG(duration_hours), 1) as avg_hours
  FROM "cnc.machine_downtime.Downtime"
  GROUP BY quarter, downtime_type
  ORDER BY quarter, total_hours DESC;

Step 2: "Most problematic machines"
→ SELECT
    d.machine_id,
    m.manufacturer_model,
    COUNT(*) as downtime_events,
    SUM(d.duration_hours) as total_downtime_hours,
    SUM(d.repair_cost_rm) as total_repair_cost_rm
  FROM "cnc.machine_downtime.Downtime" d
  JOIN "cnc.machines.Equipment" m ON d.machine_id = m.machine_id
  GROUP BY d.machine_id, m.manufacturer_model
  ORDER BY total_downtime_hours DESC
  LIMIT 10;

Step 3: "Shift to preventive maintenance in Q3-Q4?"
→ SELECT
    quarter,
    SUM(CASE WHEN downtime_type = 'Scheduled Maintenance' THEN duration_hours ELSE 0 END) as scheduled_hours,
    SUM(CASE WHEN downtime_type = 'Unscheduled Breakdown' THEN duration_hours ELSE 0 END) as breakdown_hours,
    ROUND(SUM(CASE WHEN downtime_type = 'Unscheduled Breakdown' THEN duration_hours ELSE 0 END) /
          SUM(CASE WHEN downtime_type = 'Scheduled Maintenance' THEN duration_hours ELSE 0 END), 2) as breakdown_to_preventive_ratio
  FROM "cnc.machine_downtime.Downtime"
  GROUP BY quarter
  ORDER BY quarter;

Step 4: "Schedule adherence trending"
→ SELECT
    quarter,
    ROUND(AVG(schedule_adherence_pct), 1) as avg_adherence_pct,
    ROUND(AVG(variance_hours), 1) as avg_variance_hours
  FROM "cnc.production_schedule.Schedule"
  GROUP BY quarter
  ORDER BY quarter;

Step 5: "Most common delay reasons"
→ SELECT
    delay_reason,
    COUNT(*) as occurrences,
    ROUND(AVG(variance_hours), 1) as avg_delay_hours
  FROM "cnc.production_schedule.Schedule"
  WHERE delay_reason IS NOT NULL
  GROUP BY delay_reason
  ORDER BY occurrences DESC;
```

**Key Insight:** Q1-Q2 reactive maintenance causes 2.5x more downtime than preventive. Q3-Q4 improvement visible.

---

### Sequence 5: Quality Performance Trending

**Objective:** Track quality improvements from Q3 interventions

```
Step 1: "First article pass rate by quarter"
→ SELECT
    j.quarter,
    COUNT(DISTINCT p.program_id) as total_programs,
    SUM(CASE WHEN p.first_article_pass THEN 1 ELSE 0 END) as passed_first_time,
    ROUND(SUM(CASE WHEN p.first_article_pass THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as pass_rate_pct
  FROM "cnc.program_validation.Programs" p
  JOIN "cnc.job_orders.Orders" j ON p.job_id = j.job_id
  GROUP BY j.quarter
  ORDER BY j.quarter;

Step 2: "Inspection failure rate by inspector"
→ SELECT
    inspector_name,
    inspection_type,
    COUNT(*) as inspections,
    SUM(CASE WHEN within_tolerance THEN 1 ELSE 0 END) as passed,
    ROUND((COUNT(*) - SUM(CASE WHEN within_tolerance THEN 1 ELSE 0 END)) * 100.0 / COUNT(*), 1) as failure_rate_pct
  FROM "cnc.quality_inspections.Inspections"
  GROUP BY inspector_name, inspection_type
  HAVING COUNT(*) > 50
  ORDER BY failure_rate_pct DESC;

Step 3: "Operator performance - who causes the most rework?"
→ SELECT
    operator_name,
    operator_skill,
    COUNT(*) as executions,
    SUM(CASE WHEN prove_out_result IN ('Fail', 'Rework') THEN 1 ELSE 0 END) as failed_runs,
    ROUND(SUM(CASE WHEN prove_out_result IN ('Fail', 'Rework') THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as failure_pct,
    ROUND(AVG(total_hours), 2) as avg_hours_per_job
  FROM "cnc.job_execution.Execution"
  GROUP BY operator_name, operator_skill
  HAVING COUNT(*) > 10
  ORDER BY failure_pct DESC;

Step 4: "Quality by complexity and material"
→ SELECT
    j.complexity,
    j.material_type,
    COUNT(*) as jobs,
    AVG(p.prove_out_attempts) as avg_attempts,
    ROUND(SUM(CASE WHEN j.quantity_scrapped > 0 THEN 1 ELSE 0 END) * 100.0 / COUNT(*), 1) as scrap_rate_pct
  FROM "cnc.job_orders.Orders" j
  JOIN "cnc.program_validation.Programs" p ON j.job_id = p.job_id
  GROUP BY j.complexity, j.material_type
  HAVING COUNT(*) > 5
  ORDER BY avg_attempts DESC;
```

**Key Insight:** Pass rate improved from 82% (Q1) to 85% (Q4), but still far below world-class 95%+.

---

### Sequence 6: Customer Churn Analysis

**Objective:** Identify why customers are leaving

```
Step 1: "Which customers were lost?"
→ SELECT
    customer_name,
    industry,
    total_orders,
    quality_incidents,
    on_time_delivery_pct,
    status
  FROM "cnc.customer_orders.Customers"
  WHERE status = 'Lost'
  ORDER BY total_orders DESC;

Step 2: "Correlation between quality issues and customer retention"
→ SELECT
    CASE
      WHEN quality_incidents = 0 THEN '0 incidents'
      WHEN quality_incidents <= 2 THEN '1-2 incidents'
      ELSE '3+ incidents'
    END as incident_bucket,
    COUNT(*) as customers,
    SUM(CASE WHEN status = 'Lost' THEN 1 ELSE 0 END) as lost_customers,
    ROUND(AVG(on_time_delivery_pct), 1) as avg_otd_pct
  FROM "cnc.customer_orders.Customers"
  WHERE total_orders > 0
  GROUP BY incident_bucket;

Step 3: "On-time delivery performance by customer priority"
→ SELECT
    priority,
    COUNT(*) as customers,
    AVG(total_orders) as avg_orders,
    AVG(on_time_delivery_pct) as avg_otd_pct,
    AVG(quality_incidents) as avg_quality_incidents
  FROM "cnc.customer_orders.Customers"
  WHERE total_orders > 0
  GROUP BY priority;

Step 4: "New customer acquisition"
→ SELECT
    customer_name,
    industry,
    q3_orders,
    q4_orders,
    total_orders
  FROM "cnc.customer_orders.Customers"
  WHERE q1_orders = 0 AND q2_orders = 0 AND (q3_orders > 0 OR q4_orders > 0);
```

**Key Insight:** Lost 5 customers (all had 3+ quality incidents, <70% OTD), gained 7 new ones in Q3-Q4.

---

## Query Examples

### Job Performance Overview
```sql
SELECT
    j.job_id,
    j.customer_name,
    j.part_type,
    j.complexity_score,
    j.quantity_ordered,
    j.on_time,
    p.prove_out_attempts,
    c.cost_variance_pct,
    c.gross_margin_pct
FROM "cnc.job_orders.Orders" j
JOIN "cnc.program_validation.Programs" p ON j.job_id = p.job_id
JOIN "cnc.cost_analysis.Job_Costing" c ON j.job_id = c.job_id
WHERE j.status = 'Shipped'
ORDER BY c.gross_margin_pct ASC
LIMIT 20;
```

### Material Usage by Quarter
```sql
SELECT
    SUBSTR(week_starting, 1, 7) as month,
    material_type,
    SUM(issued_to_jobs_kg) as total_issued_kg,
    AVG(unit_cost_rm) as avg_cost_per_kg,
    SUM(issued_to_jobs_kg * unit_cost_rm) as total_cost_rm
FROM "cnc.material_inventory.Inventory"
GROUP BY month, material_type
ORDER BY month, total_cost_rm DESC;
```

### Programmer Efficiency
```sql
SELECT
    programmer_name,
    programmer_skill,
    COUNT(*) as programs,
    AVG(programming_hours) as avg_hours,
    AVG(prove_out_attempts) as avg_attempts,
    SUM(CASE WHEN first_article_pass THEN 1 ELSE 0 END) * 100.0 / COUNT(*) as first_pass_pct
FROM "cnc.program_validation.Programs"
GROUP BY programmer_name, programmer_skill
ORDER BY first_pass_pct DESC;
```

### Cross-Table: Job Profitability with Quality Metrics
```sql
SELECT
    j.job_id,
    j.customer_name,
    j.complexity,
    p.prove_out_attempts,
    CASE WHEN s.scrap_id IS NOT NULL THEN 'Yes' ELSE 'No' END as had_scrap,
    c.cost_variance_pct,
    c.gross_margin_pct
FROM "cnc.job_orders.Orders" j
JOIN "cnc.program_validation.Programs" p ON j.job_id = p.job_id
JOIN "cnc.cost_analysis.Job_Costing" c ON j.job_id = c.job_id
LEFT JOIN "cnc.scrap_rework.Scrap" s ON j.job_id = s.job_id
WHERE c.gross_margin_pct IS NOT NULL
ORDER BY c.gross_margin_pct ASC;
```

## Data Features Demonstrated

### Operational Complexity
- **Multi-attempt workflows**: Prove-out failures requiring 2-4 iterations
- **Time-series tracking**: Weekly inventory, schedule adherence over 52 weeks
- **Quality measurements**: 3-8 dimensional checks per part with tolerances
- **Cost accumulation**: Material, labor, machine, tooling, overhead rollups

### Business Intelligence
- **Root cause analysis**: Scrap reasons, downtime causes, error types
- **Skill-based variance**: Senior vs junior programmer/operator performance
- **Quarterly trending**: Failure rates, costs, OEE improving Q3-Q4
- **Customer behavior**: Churn correlation with quality issues, payment terms

### Data Types
- **Decimals**: Dimensional measurements (±0.01mm tolerances)
- **Dates**: Order, programming, execution, inspection, delivery dates
- **Booleans**: on_time, preventable, simulation_performed
- **Enums**: Status, complexity, shift, machine_type

### Manufacturing Realism
- **Pareto distribution**: 20% customers = 60% revenue
- **Weibull failures**: Tool life, machine breakdowns
- **Learning curves**: Repeat jobs faster, skilled operators more consistent
- **Seasonal patterns**: Material costs, customer ordering

## Root Cause: Manual Validation Gap

The data reveals Tech Holdings' core problem:

**No automated G-code validation** → Errors caught only at expensive prove-out stage → RM500-5000 per failure → 18-21% failure rate → Labor costs 30% over estimate → Negative margins on complex jobs → Customer churn → Revenue decline

**Q3-Q4 Improvement Attempts:**
- Mandatory CAM simulation (75% adoption)
- Senior programmer review before prove-out
- Preventive maintenance schedule

**Result:** Partial improvement (15% failure rate, 10% scrap) but still unprofitable. Manual processes cannot achieve world-class 95%+ first-pass yield.

**Bottom Line:** Without digital validation tools, Tech remains uncompetitive despite process improvements.
