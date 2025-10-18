# Finance Examples: Kopitiam Kita Sdn Bhd

Real-world financial data for **Kopitiam Kita Sdn Bhd**, a Malaysian coffeehouse chain with outlets across Kuala Lumpur, Penang, Johor Bahru, Kota Kinabalu, and Kuching.

**Business Context:**
- Traditional Malaysian kopitiam (coffeehouse) chain
- Products: Kopi Susu, Teh Tarik, Kaya Toast, Roti Bakar, Nasi Lemak
- Services: Walk-in customers, corporate catering, events
- Fiscal Year: 2024 (All amounts in MYR - Malaysian Ringgit)

## Files

| File | Size | Rows | Use Case |
|------|------|------|----------|
| `general_ledger.xlsx` | 58KB | 1,000 | Journal entries, general ledger transactions |
| `financial_statements.xlsx` | 7KB | 3 sheets | P&L, Balance Sheet, Cash Flow statements |
| `accounts_receivable.xlsx` | 22KB | 300 | AR aging, invoice tracking, collections |
| `revenue_by_segment.xlsx` | 60KB | 1,020 | Revenue analysis by product/region/segment |
| `budget_vs_actuals.xlsx` | 8KB | 60 | Budget variance analysis by department |
| `invoice_register.xlsx` | 36KB | 500 | Invoice tracking, payment status |
| `trial_balance.xlsx` | 6KB | 24 | Chart of accounts, trial balance (messy headers) |
| `cash_flow_forecast.xlsx` | 6KB | 9 | 12-month cash flow projection (wide format) |
| `expense_reports.xlsx` | 15KB | 200 | Employee expense submissions |
| `financial_ratios.xlsx` | 5KB | 13 | Quarterly financial metrics and KPIs |

**Total: ~3,100 financial records across 10 files**

## Generate Examples

```bash
python examples/create_finance_examples.py
```

## Usage

```bash
# Load with proper headers and type hints
mcp-excel --path examples --alias finance --overrides examples/finance_overrides.yaml --watch
```

## Prompt Chain Sequences

Interactive analysis journeys demonstrating how to explore Kopitiam Kita's financial data step-by-step.

### Sequence 1: General Ledger Deep Dive → MYR 12M+ Total Debits

**Objective:** Understand the company's transaction volume and validate accounting entries

```
Step 1: "Load the finance examples with proper type hints"
→ mcp-excel --path examples --alias finance --overrides examples/finance_overrides.yaml

Step 2: "Show me all available tables"
→ SELECT * FROM "finance.__tables";

Step 3: "How many journal entries do we have in the general ledger?"
→ SELECT COUNT(*) as total_entries FROM "finance.general_ledger.entries";
Result: 1,000 entries

Step 4: "What's the total debit amount across all entries?"
→ SELECT SUM(COALESCE(debit, 0)) as total_debits
  FROM "finance.general_ledger.entries";
Result: MYR 12,024,765.31

Step 5: "Verify debits equal credits (accounting balance check)"
→ SELECT
    SUM(COALESCE(debit, 0)) as total_debits,
    SUM(COALESCE(credit, 0)) as total_credits,
    SUM(COALESCE(debit, 0)) - SUM(COALESCE(credit, 0)) as difference
  FROM "finance.general_ledger.entries";

Step 6: "Which accounts have the most activity?"
→ SELECT account_name, COUNT(*) as transactions
  FROM "finance.general_ledger.entries"
  GROUP BY account_name
  ORDER BY transactions DESC
  LIMIT 10;
```

**Key Insight:** 1,000 transactions totaling MYR 12M+ in debits, balanced accounting entries confirmed.

---

### Sequence 2: AR Aging Analysis → MYR 11.1M Outstanding

**Objective:** Identify collection risks and prioritize follow-ups

```
Step 1: "What's our total accounts receivable balance?"
→ SELECT SUM(balance) as total_ar
  FROM "finance.accounts_receivable.ar_aging"
  WHERE balance > 0;
Result: MYR 11,100,915.88

Step 2: "Break down AR by aging buckets"
→ SELECT
    aging_bucket,
    COUNT(*) as invoice_count,
    SUM(balance) as total_outstanding,
    ROUND(SUM(balance) * 100.0 / (SELECT SUM(balance)
      FROM "finance.accounts_receivable.ar_aging"
      WHERE balance > 0), 2) as pct_of_total
  FROM "finance.accounts_receivable.ar_aging"
  WHERE balance > 0
  GROUP BY aging_bucket
  ORDER BY CASE aging_bucket
    WHEN 'Current' THEN 1
    WHEN '31-60 days' THEN 2
    WHEN '61-90 days' THEN 3
    ELSE 4 END;

Step 3: "Who are our top 10 customers by outstanding balance?"
→ SELECT
    customer_name,
    COUNT(*) as open_invoices,
    SUM(balance) as total_due,
    MAX(days_outstanding) as oldest_invoice_days
  FROM "finance.accounts_receivable.ar_aging"
  WHERE balance > 0
  GROUP BY customer_name
  ORDER BY total_due DESC
  LIMIT 10;

Step 4: "Focus on high-risk: Over 90 days past due"
→ SELECT customer_name, invoice_number, invoice_date, balance
  FROM "finance.accounts_receivable.ar_aging"
  WHERE aging_bucket = '90+ days' AND balance > 0
  ORDER BY balance DESC;
```

**Key Insight:** MYR 11.1M outstanding, need to prioritize customers in 90+ days bucket for collections.

---

### Sequence 3: Revenue & Margin Analysis → MYR 257M Revenue, 20.6% Margin

**Objective:** Understand revenue drivers and profitability by segment

```
Step 1: "What's our total revenue for the year?"
→ SELECT SUM(revenue) as total_revenue
  FROM "finance.revenue_by_segment.revenue";
Result: MYR 257,675,901.23

Step 2: "What's our average gross margin percentage?"
→ SELECT AVG(margin_pct) as avg_margin
  FROM "finance.revenue_by_segment.revenue";
Result: 20.6%

Step 3: "Which regions are most profitable?"
→ SELECT
    region,
    SUM(revenue) as total_revenue,
    SUM(gross_profit) as total_profit,
    AVG(margin_pct) as avg_margin_pct
  FROM "finance.revenue_by_segment.revenue"
  GROUP BY region
  ORDER BY total_revenue DESC;

Step 4: "Which products have the highest margins?"
→ SELECT
    product,
    SUM(revenue) as revenue,
    AVG(margin_pct) as avg_margin
  FROM "finance.revenue_by_segment.revenue"
  GROUP BY product
  ORDER BY avg_margin DESC;

Step 5: "Monthly revenue trend - are we growing?"
→ SELECT
    month,
    SUM(revenue) as monthly_revenue,
    SUM(gross_profit) as monthly_profit
  FROM "finance.revenue_by_segment.revenue"
  GROUP BY month
  ORDER BY month;

Step 6: "Segment performance: Corporate vs Walk-in vs Catering"
→ SELECT
    segment,
    SUM(revenue) as revenue,
    COUNT(DISTINCT month) as active_months,
    SUM(revenue) / COUNT(DISTINCT month) as avg_monthly_revenue
  FROM "finance.revenue_by_segment.revenue"
  GROUP BY segment
  ORDER BY revenue DESC;
```

**Key Insight:** MYR 257M revenue with 20.6% average margin. Kuala Lumpur & Selangor region driving majority of sales.

---

### Sequence 4: Invoice Tracking → 318 Paid Invoices, MYR 13.1M

**Objective:** Monitor payment collection and identify payment patterns

```
Step 1: "How many invoices have been paid?"
→ SELECT COUNT(*) as paid_count
  FROM "finance.invoice_register.invoices"
  WHERE status = 'Paid';
Result: 318 invoices

Step 2: "What's the total value of paid invoices?"
→ SELECT SUM(total_amount) as total_paid
  FROM "finance.invoice_register.invoices"
  WHERE status = 'Paid';
Result: MYR 13,148,737.96

Step 3: "Payment status breakdown"
→ SELECT
    status,
    COUNT(*) as invoice_count,
    SUM(total_amount) as total_value,
    ROUND(AVG(total_amount), 2) as avg_invoice_value
  FROM "finance.invoice_register.invoices"
  GROUP BY status
  ORDER BY total_value DESC;

Step 4: "What payment methods do customers prefer?"
→ SELECT
    payment_method,
    COUNT(*) as count,
    SUM(total_amount) as total_value
  FROM "finance.invoice_register.invoices"
  WHERE status = 'Paid'
  GROUP BY payment_method
  ORDER BY total_value DESC;

Step 5: "Average days to payment (for paid invoices)"
→ SELECT
    AVG(JULIANDAY(payment_date) - JULIANDAY(invoice_date)) as avg_days_to_pay
  FROM "finance.invoice_register.invoices"
  WHERE status = 'Paid' AND payment_date IS NOT NULL;

Step 6: "Identify overdue invoices for follow-up"
→ SELECT invoice_id, customer_name, due_date, total_amount
  FROM "finance.invoice_register.invoices"
  WHERE status = 'Overdue'
  ORDER BY total_amount DESC;
```

**Key Insight:** 318 paid invoices totaling MYR 13.1M. Wire Transfer is the most common payment method.

---

### Sequence 5: Budget Variance Analysis → MYR 4.2M Over-Budget

**Objective:** Identify cost overruns and departments needing attention

```
Step 1: "Which departments are over budget?"
→ SELECT
    department,
    SUM(budget) as total_budget,
    SUM(actual) as total_actual,
    SUM(variance) as total_variance
  FROM "finance.budget_vs_actuals.analysis"
  WHERE variance > 0
  GROUP BY department
  ORDER BY total_variance DESC;

Step 2: "Total over-budget spending across the company"
→ SELECT SUM(actual) as total_overspend
  FROM "finance.budget_vs_actuals.analysis"
  WHERE variance > 0;
Result: MYR 4,228,972.67

Step 3: "Which expense categories are problematic?"
→ SELECT
    expense_category,
    COUNT(*) as dept_count,
    SUM(variance) as total_overrun,
    AVG(variance_pct) as avg_variance_pct
  FROM "finance.budget_vs_actuals.analysis"
  WHERE variance > 0
  GROUP BY expense_category
  ORDER BY total_overrun DESC
  LIMIT 10;

Step 4: "Overall budget utilization rate"
→ SELECT
    SUM(budget) as total_budget,
    SUM(actual) as total_actual,
    ROUND(SUM(actual) * 100.0 / SUM(budget), 2) as utilization_pct
  FROM "finance.budget_vs_actuals.analysis";

Step 5: "Find departments under budget (potential savings)"
→ SELECT
    department,
    expense_category,
    budget,
    actual,
    variance,
    variance_pct
  FROM "finance.budget_vs_actuals.analysis"
  WHERE variance < 0
  ORDER BY variance ASC
  LIMIT 10;

Step 6: "Outlet Operations deep dive (if over budget)"
→ SELECT
    expense_category,
    budget,
    actual,
    variance,
    variance_pct
  FROM "finance.budget_vs_actuals.analysis"
  WHERE department = 'Outlet Operations'
  ORDER BY ABS(variance) DESC;
```

**Key Insight:** MYR 4.2M over-budget spending. Need to investigate "Raw Materials" and "Marketing & Advertising" categories.

---

## Query Examples

### General Ledger Analysis
```sql
-- Monthly journal entry summary
SELECT
    STRFTIME('%Y-%m', date) as month,
    account_name,
    SUM(COALESCE(debit, 0)) as total_debits,
    SUM(COALESCE(credit, 0)) as total_credits
FROM "finance.general_ledger.entries"
GROUP BY month, account_name
ORDER BY month DESC, account_name;

-- Account balance verification
SELECT
    account_name,
    SUM(COALESCE(debit, 0)) - SUM(COALESCE(credit, 0)) as balance
FROM "finance.general_ledger.entries"
GROUP BY account_name
ORDER BY ABS(balance) DESC;
```

### AR Aging Analysis
```sql
-- AR aging summary
SELECT
    aging_bucket,
    COUNT(*) as invoice_count,
    SUM(balance) as total_outstanding
FROM "finance.accounts_receivable.ar_aging"
WHERE balance > 0
GROUP BY aging_bucket
ORDER BY
    CASE aging_bucket
        WHEN 'Current' THEN 1
        WHEN '31-60 days' THEN 2
        WHEN '61-90 days' THEN 3
        ELSE 4
    END;

-- Top 10 customers by outstanding balance
SELECT
    customer_name,
    COUNT(*) as open_invoices,
    SUM(balance) as total_due
FROM "finance.accounts_receivable.ar_aging"
WHERE balance > 0
GROUP BY customer_name
ORDER BY total_due DESC
LIMIT 10;
```

### Revenue Analysis
```sql
-- Revenue by segment and region
SELECT
    segment,
    region,
    SUM(revenue) as total_revenue,
    SUM(gross_profit) as total_profit,
    AVG(margin_pct) as avg_margin
FROM "finance.revenue_by_segment.revenue"
GROUP BY segment, region
ORDER BY total_revenue DESC;

-- Monthly revenue trend
SELECT
    month,
    SUM(revenue) as monthly_revenue,
    SUM(gross_profit) as monthly_profit
FROM "finance.revenue_by_segment.revenue"
GROUP BY month
ORDER BY month;
```

### Budget Variance
```sql
-- Department budget performance
SELECT
    department,
    SUM(budget) as total_budget,
    SUM(actual) as total_actual,
    SUM(variance) as total_variance,
    ROUND(SUM(actual) / SUM(budget) * 100, 2) as budget_utilization_pct
FROM "finance.budget_vs_actuals.analysis"
GROUP BY department
ORDER BY total_variance DESC;

-- Over-budget expense categories
SELECT
    expense_category,
    COUNT(*) as dept_count,
    SUM(variance) as total_overrun
FROM "finance.budget_vs_actuals.analysis"
WHERE variance > 0
GROUP BY expense_category
ORDER BY total_overrun DESC;
```

### Financial Statements
```sql
-- Income statement quarterly comparison
SELECT * FROM "finance.financial_statements.income_statement"
ORDER BY line_item;

-- Balance sheet analysis
SELECT
    category,
    SUM(amount) as total
FROM "finance.financial_statements.balance_sheet"
WHERE category IN ('Asset', 'Liability', 'Equity')
GROUP BY category;

-- Cash flow by section
SELECT
    section,
    SUM(amount) as net_cash
FROM "finance.financial_statements.cash_flow"
WHERE section IN ('Operating', 'Investing', 'Financing')
GROUP BY section;
```

### Invoice Analysis
```sql
-- Payment status summary
SELECT
    status,
    COUNT(*) as invoice_count,
    SUM(total_amount) as total_value
FROM "finance.invoice_register.invoices"
GROUP BY status;

-- Payment method breakdown (paid invoices only)
SELECT
    payment_method,
    COUNT(*) as count,
    SUM(total_amount) as total
FROM "finance.invoice_register.invoices"
WHERE status = 'Paid'
GROUP BY payment_method;
```

### Expense Reports
```sql
-- Department expense summary
SELECT
    department,
    COUNT(*) as report_count,
    SUM(amount) as total_expenses,
    AVG(amount) as avg_expense
FROM "finance.expense_reports.reports"
WHERE status = 'Approved'
GROUP BY department
ORDER BY total_expenses DESC;

-- Expense type analysis
SELECT
    expense_type,
    COUNT(*) as count,
    SUM(amount) as total,
    MIN(amount) as min,
    MAX(amount) as max
FROM "finance.expense_reports.reports"
GROUP BY expense_type
ORDER BY total DESC;
```

### Financial Ratios & KPIs
```sql
-- Quarterly trend analysis
SELECT * FROM "finance.financial_ratios.quarterly_metrics"
WHERE metric IN ('Net Profit Margin %', 'Return on Equity %', 'Revenue Growth % YoY');

-- Key metrics comparison
SELECT
    metric,
    "Q1 2024" as q1,
    "Q4 2024" as q4,
    "Q4 2024" - "Q1 2024" as improvement
FROM "finance.financial_ratios.quarterly_metrics"
WHERE metric LIKE '%Margin%' OR metric LIKE '%Return%';
```

### Cross-Table Analysis
```sql
-- Revenue vs AR correlation
SELECT
    r.month,
    SUM(r.revenue) as monthly_revenue,
    (SELECT SUM(balance) FROM "finance.accounts_receivable.ar_aging"
     WHERE STRFTIME('%Y-%m', invoice_date) = r.month) as ar_balance
FROM "finance.revenue_by_segment.revenue" r
GROUP BY r.month
ORDER BY r.month;
```

## Data Features Demonstrated

### Headers & Formatting
- **trial_balance.xlsx**: Multi-row headers with metadata (skip_rows, skip_footer)
- **cash_flow_forecast.xlsx**: Wide format with 12 month columns
- **financial_statements.xlsx**: Multi-sheet workbook with related data

### Data Types
- **Decimals**: All monetary amounts with proper precision
- **Dates**: Various date formats (ISO, invoice dates, payment dates)
- **Nullables**: Missing payment dates, optional fields

### Business Logic
- **Calculated columns**: Balance = Amount - Paid, Variance = Actual - Budget
- **Aging buckets**: Current, 31-60, 61-90, 90+ days
- **Percentages**: Margins, variance %, utilization rates

### Data Quality
- **Completeness**: Mix of complete and incomplete records
- **Validation**: Debit/credit balance checks, date logic
- **Referential integrity**: Customer names across AR and invoices
