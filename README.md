# 📊 N Level CPA Excel Tutorials

A comprehensive collection of Excel practice materials designed for N Level Computer Applications (CPA) students. This repository contains interactive workbooks, detailed guides, and automated workbook generators to help students master essential Excel skills.

## 🎯 About

This repository was created by **Ernest Tan** using the [Excel Study Coach](https://chatgpt.com/g/g-68cf55b2147881919b2b4acfbdb427e3-excel-study-coach) GPT to provide structured, hands-on Excel learning materials specifically tailored for N Level CPA curriculum requirements.

## 📚 Tutorial Topics

### Core Excel Skills
- **Core Functions** - Master SUM, AVERAGE, MIN, MAX, COUNT, COUNTA
- **IF Functions** - Learn logical decision-making in Excel
- **Text Functions** - String manipulation and text processing
- **Date & Time** - Working with dates, times, and calculations
- **Lookup Functions** - VLOOKUP, HLOOKUP, and advanced lookup techniques

### Data Analysis & Visualization
- **Conditional Formatting** - Visual data analysis and highlighting
- **Charts & Graphs** - Creating effective data visualizations
- **Sorting & Filtering** - Organizing and analyzing data sets
- **Simple Data Analysis** - Basic statistical analysis techniques

### Advanced Functions
- **COUNTIFS Functions** - Multi-criteria counting and analysis
- **N Level COUNTIFS** - Advanced conditional counting for exam preparation

## 📁 Repository Structure

```
exceltutor/
├── README.md                           # This file
├── venv/                              # Python virtual environment
│
├── Practice Materials/
│   ├── Charts_Practice.xlsx           # Interactive chart creation exercises
│   ├── Charts_Practice.pdf            # Step-by-step chart tutorial
│   ├── Conditional_Formatting_Practice.xlsx
│   ├── Conditional_Formatting_Practice.pdf
│   ├── Core_Functions_Practice.xlsx   # Basic Excel functions
│   ├── Core_Functions_Practice.pdf
│   ├── dates_time_practice.xlsx       # Date/time calculations
│   ├── dates_time_practice.pdf
│   ├── IF_Function_Starter.xlsx       # Logical functions
│   ├── IF_Function_Starter.pdf
│   ├── lookup_practice.xlsx           # VLOOKUP and lookup functions
│   ├── lookup_practice.pdf
│   ├── NLevel_COUNTIFS_Practice.xlsx  # Advanced counting functions
│   ├── NLevel_COUNTIFS_Practice.pdf
│   ├── Simple_Data_Analysis_Starter.xlsx
│   ├── Simple_Data_Analysis_Starter.pdf
│   ├── Sorting_Filtering_Practice.xlsx
│   ├── Sorting_Filtering_Practice.pdf
│   ├── Text_Functions_Practice.xlsx   # String manipulation
│   └── Text_Functions_Practice.pdf
│
└── Generators/                        # Python scripts to create workbooks
    ├── topic4.py                      # Core Functions generator
    ├── topic5.py                      # IF Functions generator
    ├── topic6.py                      # Text Functions generator
    ├── topic7.py                      # Date/Time generator
    ├── topic8.py                      # Lookup Functions generator
    ├── topic9.py                      # Conditional Formatting generator
    ├── topic10a.py                    # Sorting & Filtering generator
    ├── topic10b.py                    # Charts generator
    ├── topic10c.py                    # Data Analysis generator
    └── topic11.py                     # COUNTIFS generator
```

## 🚀 Getting Started

### For Students

1. **Download the Excel files** (.xlsx) for hands-on practice
2. **Read the PDF guides** for step-by-step instructions
3. **Follow the structured approach** in each workbook:
   - Start with the **Instructions** sheet
   - Review sample data on the **Data** sheet
   - Complete exercises on the **Tasks** sheet
   - Use **Hints** when stuck
   - Check your work with the **Answers** sheet
   - Track progress on the **Checklist** sheet
   - Reference the **Lookup** sheet for quick function syntax

### For Educators

1. **Use the pre-built workbooks** for classroom activities
2. **Customize content** by running the Python generators
3. **Track student progress** using the built-in checklists
4. **Adapt difficulty levels** by modifying the generator scripts

## 🛠️ Technical Setup (For Customization)

If you want to modify or regenerate the workbooks:

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Installation
```bash
# Clone the repository
git clone <repository-url>
cd exceltutor

# Create and activate virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install openpyxl
```

### Generating Workbooks
```bash
# Run any generator script
python topic4.py          # Creates Core_Functions_Practice.xlsx
python topic10a.py        # Creates Sorting_Filtering_Practice.xlsx
# ... and so on
```

## 📖 How Each Workbook Works

Every practice workbook follows a consistent 6-sheet structure:

1. **Instructions** 📋 - Learning objectives and overview
2. **Data** 📊 - Sample datasets for practice
3. **Tasks** ✏️ - Exercises with input cells highlighted in yellow
4. **Hints** 💡 - Gentle guidance when students get stuck
5. **Answers** ✅ - Complete solutions for self-checking
6. **Checklist** ☑️ - Skills tracking and progress monitoring
7. **Lookup** 📚 - Quick reference for function syntax

## 🎓 Learning Approach

### Progressive Difficulty
- Start with basic functions and build complexity
- Each topic builds on previous knowledge
- Real-world examples and scenarios

### Self-Paced Learning
- Clear instructions and objectives
- Built-in hints system
- Self-assessment tools

### Practical Application
- Hands-on exercises with immediate feedback
- Real business scenarios and data
- Industry-standard Excel techniques

## 🤝 Contributing

This repository is designed for educational use. If you're an educator or student with suggestions for improvements:

1. Fork the repository
2. Create your feature branch
3. Make your changes
4. Submit a pull request

## 📞 Support

For questions about using these materials:
- Check the **Hints** and **Lookup** sheets in each workbook
- Review the accompanying PDF guides
- Consult the [Excel Study Coach](https://chatgpt.com/g/g-68cf55b2147881919b2b4acfbdb427e3-excel-study-coach) GPT for additional help

## 📄 License

This educational content is provided for learning purposes. Please respect copyright and use responsibly in educational settings.

## 🙏 Acknowledgments

- Created with assistance from ChatGPT's Excel Study Coach
- Designed specifically for N Level CPA curriculum
- Built with Python and the openpyxl library

---

**Happy Learning! 📈✨**

*Master Excel one function at a time with structured, hands-on practice materials designed for success.*
