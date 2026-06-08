# Amazon EFS & AWS Backup Cost Optimizer

This tool evaluates **Amazon EFS file systems** and **AWS Backup recovery points** for cost optimization opportunities. It generates a colour-coded Excel report with actionable recommendations and a savings summary.

## What It Checks

### EFS Cost Optimization

| # | Check | Issue Detected | Recommendation |
|---|-------|----------------|----------------|
| 1 | **No Mount Targets** | File system has no mount points — likely unused | Delete filesystem or investigate |
| 2 | **Missing Lifecycle Policy** | No IA transition configured | Add lifecycle policy (save up to 92% on infrequent data) |
| 3 | **Size Constant 90 Days** | No data written in 90 days — possibly abandoned | Review usage, consider deletion |
| 4 | **No Backup Configured** | Neither EFS automatic backup nor AWS Backup protection | Enable backup or document risk acceptance |

### AWS Backup Cost Optimization

| # | Check | Issue Detected | Recommendation |
|---|-------|----------------|----------------|
| 1 | **Infinite Retention** | Recovery point set to NEVER expire | Set retention policy to avoid unbounded cost growth |
| 2 | **Exceeds Max Retention** | Recovery point older than specified max retention | Delete to reclaim storage costs |
| 3 | **Expired Status** | Recovery point in EXPIRED state | Delete — no longer needed |

## Decision Logic

### EFS Recommendations

```
┌──────────────────────────────┐
│ For each EFS File System     │
└──────────────┬───────────────┘
               │
               ▼
┌──────────────────────────────┐   YES   ┌────────────────────────────────┐
│ Has mount targets?           │────────▶│ Check lifecycle & size...       │
└──────────────┬───────────────┘         └────────────────────────────────┘
               │ NO
               ▼
┌────────────────────────────────────┐
│ 🔴 REVIEW — NO MOUNT TARGETS      │
│ (filesystem may be unused)         │
└────────────────────────────────────┘
```

### Backup Recommendations

```
┌──────────────────────────────────┐
│ For each Recovery Point          │
└──────────────┬───────────────────┘
               │
               ▼
┌──────────────────────────────┐   YES   ┌───────────────────────────────┐
│ Status = EXPIRED?            │────────▶│ 🔴 DELETE — EXPIRED            │
└──────────────┬───────────────┘         └───────────────────────────────┘
               │ NO
               ▼
┌──────────────────────────────┐   YES   ┌───────────────────────────────┐
│ Age > max retention?         │────────▶│ 🟡 DELETE — EXCEEDS RETENTION  │
└──────────────┬───────────────┘         └───────────────────────────────┘
               │ NO
               ▼
┌──────────────────────────────┐   YES   ┌───────────────────────────────┐
│ Retention = NEVER?           │────────▶│ 🟠 REVIEW — SET RETENTION      │
└──────────────┬───────────────┘         └───────────────────────────────┘
               │ NO
               ▼
┌───────────────────────────────┐
│ 🟢 OK — No action needed     │
└───────────────────────────────┘
```

## EFS Cost Context

| Storage Class | Price ($/GB-month) | Use Case |
|---------------|-------------------|----------|
| EFS Standard | $0.30 | Frequently accessed data |
| EFS Infrequent Access (IA) | $0.025 | Data accessed < once/month |
| EFS Archive | $0.008 | Data accessed few times/year |

**Lifecycle policies can save up to 92%** by automatically moving infrequent data to IA/Archive tiers.

## AWS Backup Cost Context

| Storage | Price ($/GB-month) | Notes |
|---------|-------------------|-------|
| Warm Storage | $0.05 | Immediate restore |
| Cold Storage | $0.01 | 12-hour restore time, 90-day minimum |

**Infinite retention** recovery points grow unbounded — a common cost leak in enterprise accounts.

## How to Use

### Prerequisites
- Python 3.9+
- AWS credentials configured (profile or environment variables)
- IAM permissions:
  - `elasticfilesystem:DescribeFileSystems`
  - `elasticfilesystem:DescribeMountTargets`
  - `elasticfilesystem:DescribeBackupPolicy`
  - `cloudwatch:GetMetricStatistics`
  - `backup:ListBackupVaults`
  - `backup:ListRecoveryPointsByBackupVault`
  - `backup:ListProtectedResources`

### Installation

```bash
git clone https://github.com/sharmaankush123/efs-backup-cost-optimizer.git
cd efs-backup-cost-optimizer
pip3 install -r requirements.txt
```

### Execution

```bash
# Single region
python3 efs_backup_cost_optimizer.py <aws-profile> <region>

# Multiple regions
python3 efs_backup_cost_optimizer.py <aws-profile> <region1,region2,region3>

# With max retention period (flags recovery points older than N days)
python3 efs_backup_cost_optimizer.py <aws-profile> <region(s)> <max-retention-days>

# Examples
python3 efs_backup_cost_optimizer.py default us-east-1
python3 efs_backup_cost_optimizer.py production us-east-1,eu-west-1,ap-southeast-2
python3 efs_backup_cost_optimizer.py default us-east-1,eu-west-1 365
```

### Output

The script generates an **Excel file (.xlsx)** with three sheets:

1. **EFS Cost Optimization** — per-filesystem assessment (colour-coded)
2. **AWS Backup Optimization** — per-recovery-point assessment (colour-coded)
3. **Cost Savings Summary** — total potential savings if recommendations are applied

### Colour Coding

| Colour | Meaning |
|--------|---------|
| 🟢 Green | OK — no action needed |
| 🟡 Yellow | Warning — review recommended |
| 🟠 Orange | Action needed — configure retention or lifecycle |
| 🔴 Red | Critical — unused resource or expired backup |

## Sample Output

```
╔══════════════════════════════════════════════════╗
║   EFS & AWS Backup Cost Optimizer               ║
╠══════════════════════════════════════════════════╣
║  Profile: production                            ║
║  Regions: us-east-1, eu-west-1                  ║
║  Max Retention: 365 days                        ║
╚══════════════════════════════════════════════════╝

[us-east-1] Connecting...
[us-east-1] Evaluating EFS file systems...
[us-east-1]   → 12 file systems evaluated
[us-east-1] Evaluating AWS Backup recovery points...
[us-east-1]   → 847 recovery points evaluated

Generating Excel report...
✅ Report saved: efs_backup_cost_optimization_us-east-1_eu-west-1_20260608.xlsx

Summary:
  EFS file systems evaluated: 12
  Backup recovery points evaluated: 847
  Estimated monthly savings: $2,340.00
  Estimated annual savings:  $28,080.00
```

## Common Scenarios

### Scenario 1: EFS without Lifecycle Policy
A 500 GB EFS filesystem with no lifecycle policy pays $150/month. If 80% of data is infrequently accessed, adding a lifecycle policy reduces cost to ~$40/month (**$110/month savings**).

### Scenario 2: Infinite Retention Backups
An account with 200 recovery points set to "never expire" accumulates ~$500/month in storage costs that grows every backup cycle. Setting a 90-day retention and deleting old points saves immediately.

### Scenario 3: Abandoned EFS Filesystem
A 100 GB EFS with no mount targets and constant size for 90 days costs $30/month for nothing. Deleting saves $360/year.

## Security

See [CONTRIBUTING](CONTRIBUTING.md) for more information.

## License

This library is licensed under the MIT-0 License. See the [LICENSE](LICENSE) file.
