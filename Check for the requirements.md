# Requirements

Converted from [Check for the requirements.docx](Check%20for%20the%20requirements.docx).

## 1. What is (are) the repository URL(s)?

What is (are) the repository URL(s)?

### GitHub

#### In the example Excel file:

[github_stats.xlsx](github_stats.xlsx)

#### In the code:

[fetch_and_parse_github_repo.py](utils/common/fetch_and_parse_github_repo.py)

## 2. Pre-processing and Pipeline Code

### In the example Excel file:

[results.xlsx](utils/5.1.3.pre-processing_&_pipeline_code/results.xlsx)

### In the code:

 [check_paper_appendix_for_data_preprocessing_code.py](utils/5.1.3.pre-processing_&_pipeline_code/check_paper_appendix_for_data_preprocessing_code.py)

## 3. Code Documentation Quality

What level of documentation does exist for the code?

### 3.1 README EXISTS

#### In the example Excel file:

[github_stats.xlsx](github_stats.xlsx)


#### In the code:

1. [find_readme.py](utils/common/find_readme.py) - exact extractor functions: summarize_readme_files(...), get_readme_info(...).
2. [scrape_github_data.py](utils/scrape_github_data.py) - call site: has_readme = get_readme_info(owner, repo, headers).

### 3.2 INSTALLATION INSTRUCTIONS

#### In the example Excel file:

[installation_instructions_results.xlsx](utils/5.1.4.github_code_documentation_quality/check_github_repo_installation_instructions/installation_instructions_results.xlsx)

#### In the code:

1. [check_github_repo_installation_instructions.py](utils/5.1.4.github_code_documentation_quality/check_github_repo_installation_instructions/check_github_repo_installation_instructions.py)
2. [environment_instructions_existance_check.py](utils/5.1.4.github_code_documentation_quality/environment_instructions_existance_check.py)

All the code is in [environment_instructions_existance_check.py](utils/5.1.4.github_code_documentation_quality/environment_instructions_existance_check.py) in the utils folder.
