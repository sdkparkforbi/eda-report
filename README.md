# EDA Report Generator

Automated publication-ready Exploratory Data Analysis reports with GPT-powered variable classification.

## Features

- **Automatic Variable Classification**: Uses GPT API to intelligently classify variables as ID, Time, Categorical, Numeric Discrete, or Numeric Continuous
- **Publication-Ready Figures**: Uniform-sized, high-DPI figures suitable for journal submission
- **Word Document Output**: Tables and figures formatted for direct use in manuscripts
- **Auto-Generated Manuscript Text**: Methods and Results sections drafted by GPT

## Installation

```bash
pip install git+https://github.com/yourusername/eda-report.git
```

Or clone and install locally:

```bash
git clone https://github.com/yourusername/eda-report.git
cd eda-report
pip install -e .
```

## Quick Start

```python
from eda_report import generate_report
import pandas as pd

# Load your data
df = pd.read_csv('your_data.csv')

# Generate report (outputs: eda_report.zip)
generate_report(df, download=True)
```

Output: `eda_report.zip` containing Word document + all figures

## Google Colab Usage

```python
!pip install git+https://github.com/sdkparkforbi/eda-report.git -q

from eda_report import generate_report
import pandas as pd

df = pd.read_csv('your_data.csv')
generate_report(df, download=True)
```

## Configuration

```python
result = generate_report(
    df,
    output='report.docx',
    config={
        'fig_width': 6,           # Figure width in inches
        'fig_height': 5,          # Figure height in inches
        'fig_dpi': 300,           # Figure resolution
        'n_representative': 3,    # Subjects for time series plot
        'colors': {
            'primary': '#2E86AB',
            'secondary': '#A23B72',
            'accent': '#F18F01'
        }
    },
    verbose=True,
    download=False
)
```

## Output

The generated Word document includes:

1. **Table 1**: Dataset Overview (observations, variables, subjects)
2. **Table 2**: Variable Summary Statistics (type, N, missing %, mean/SD, range)
3. **Table 3**: Categorical Variable Frequencies
4. **Figures**: Distribution plots and time series (if applicable)
5. **Draft Manuscript**: Methods and Results sections

## Variable Classification Rules

| Type | Description | Examples |
|------|-------------|----------|
| ID | Subject/group identifiers | patient_id, user_id |
| Time | Temporal indices | time_idx, date, timestamp |
| Categorical | Nominal categories (no magnitude) | gender, day_of_week |
| Numeric Discrete | Integers with meaningful magnitude | dosage, age, count |
| Numeric Continuous | Continuous measurements | glucose, temperature |

**Important**: Integer variables like medication dosage are classified as Numeric Discrete (not Categorical) because magnitude comparison is meaningful.

## Requirements

- Python >= 3.8
- pandas
- numpy
- matplotlib
- python-docx
- openai

## License

MIT License
