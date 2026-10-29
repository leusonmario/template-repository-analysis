# GitHub Template Repository Analysis

This repository contains datasets, scripts, and analysis artifacts supporting empirical research on **GitHub template repositories** across five programming languages: TypeScript, Python, JavaScript, Java, and C#.

The study investigates the domains and purposes of template repositories, their maintenance and quality characteristics, and the guidelines and pitfalls associated with their creation, maintenance, and adoption.

The research combines repository mining and quantitative and qualitative analyses. The extended study also incorporates a survey of **114 practitioners** with experience creating, maintaining, or contributing to GitHub template repositories.

## Research Questions

The study addresses three research questions:

- **RQ1 — Domains, purposes, and ownership:** What purposes and domains do GitHub template repositories serve, and what ownership patterns characterize them?
- **RQ2 — Maintenance and quality:** How are GitHub templates maintained, and how do repository characteristics and activity relate to their quality?
- **RQ3 — Guidelines and pitfalls:** What actionable insights can support practitioners in creating, maintaining, and adopting templates effectively?

The repository mining process collected **14,780 template repositories** across the five programming languages. Research-question-specific filtering and analysis procedures are documented in the manuscript.

## Repository Structure

The repository is organized into the following directories:

```text
template-repository-analysis/
├── data_collection/
├── llm_evaluation/
├── rq1/
│   ├── llm_evaluation/
│   └── owner_evaluation/
├── rq2/
│   ├── correlation_complete/
│   └── regression_complete/
├── rq3/
│   ├── llm_assistant/
│   └── rq3_samples/
└── survey/
```

### `data_collection/`

Contains scripts and datasets related to the collection and preprocessing of GitHub template repositories.

The artifacts include:

- GitHub API-based repository collection scripts.
- Repository metadata, including programming language, creation dates, and development activity.
- Language-specific datasets used in subsequent analyses.

The collection process targets public, non-forked, non-archived repositories marked as GitHub templates, updated within the preceding two years, and having at least two stars.

### `llm_evaluation/`

Contains the initial evaluation of Large Language Models (LLMs) for identifying the domains supported by GitHub template repositories.

Two models, GPT-4o-mini and DeepSeek-V3.2, were compared against a manually classified sample of 100 repositories.

The results informed the selection of the model used for the subsequent domain classification.

### `rq1/` — Domain Classification and Ownership

Contains scripts and analysis artifacts used to characterize the domains and ownership patterns of GitHub template repositories.

The analyses include:

- LLM-based classification of repositories into predefined domain categories.
- Evaluation of domain classification results.
- Analysis of repository ownership (individual users versus organizations).
- Generation of figures illustrating domain and ownership distributions.

The domain classification was conducted on repositories retained after applying the README language filter.

### `rq2/` — Maintenance and Quality Analysis

Contains scripts and datasets used to investigate maintenance and quality-related issues in GitHub template repositories.

The analyses include:

- Collection and processing of SonarQube quality metrics, including bugs, code smells, vulnerabilities, and security hotspots.
- Exploratory analysis of issue distributions and densities.
- Statistical comparisons between organization-owned and user-owned templates.
- Negative Binomial regression to investigate associations between repository characteristics and quality-related issue counts.
- Correction for multiple statistical tests using the Benjamini–Hochberg procedure.

The regression analyses account for repository size using a logarithmic KLOC offset and include repository activity, popularity, age, staleness, and ownership as explanatory variables.

The regression implementation uses HC3 robust standard errors, a fixed Negative Binomial dispersion parameter, and capping of extreme vulnerability-severity counts.

### `rq3/` — Guidelines and Pitfalls

Contains scripts and artifacts supporting the selection and assessment of template repositories for identifying actionable practices.

The analyses include:

- Construction of a composite quality score based on quality and maintenance metrics.
- Selection of repositories with contrasting quality characteristics (best, worst, and mixed).
- Organization of the repositories used for qualitative inspection.
- Exploratory LLM-based assessment of adherence to the identified guidelines and pitfalls.

The qualitative analysis examined 250 repositories across the five programming languages and identified seven guidelines and five recurring pitfalls.

An additional exploratory assessment evaluated GPT-5.2 against human judgments on 20 repositories, covering 12 adherence criteria.

### `survey/` — Practitioner Survey (Extension)

This directory contains the materials supporting the practitioner survey conducted as part of an extension of our study.

The survey collected responses from **114 practitioners** with experience creating, maintaining, or contributing to GitHub template repositories.

The survey investigates:

- **RQ1:** Motivations for creating and maintaining templates, development contexts, and intended audiences.
- **RQ2:** Maintenance practices, perceived maintenance indicators, ownership-related differences, and perceptions of quality issues propagating to generated projects.
- **RQ3:** Practitioner assessments of the proposed guidelines and pitfalls, challenges in template creation and evolution, and recommendations for improving template reuse.

The survey also investigates practitioners' use and perceptions of AI-assisted tools in template development and maintenance.

The materials in this directory support the quantitative and qualitative analyses reported in the extension, including survey responses, descriptive statistics, and thematic coding, where available in the public release.

Survey data shared through this repository are subject to anonymization and the applicable research ethics requirements.

## Requirements and Reproducibility

The scripts primarily use Python and depend on libraries such as:

- pandas
- NumPy
- statsmodels
- requests
- scikit-learn

Different analyses may require additional dependencies, including access to SonarQube or external LLM APIs.

Some scripts use relative paths to datasets and output directories. Users should inspect the relevant script and configure its dependencies and execution paths before running it.

Re-executing repository collection or LLM-based analyses may produce results that differ from the archived outputs because GitHub repositories, external APIs, and hosted models can change over time.

For reproducing the reported findings, the archived datasets and analysis outputs should be used whenever available.

## Data Availability and Responsible Use

The repository provides research materials intended to support transparency, replication, and secondary analysis.

Repository metadata originates from publicly accessible GitHub repositories. 
