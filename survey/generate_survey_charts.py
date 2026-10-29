#!/usr/bin/env python3
"""
Generate publication-quality figures for the GitHub Template Repositories
practitioner survey.

This script imports survey_statistics.py and therefore uses the exact same
data-processing and counting logic used to validate the manuscript numbers.

Input
-----
survey_answers.xlsx

Usage
-----
python generate_survey_charts.py \
    survey_answers.xlsx \
    --output_dir survey_charts

python generate_survey_charts.py survey_answers_submission.csv --output_dir survey_charts_test

Outputs
-------
survey_charts/
    data_collection/
    rq1/
    rq2/
    rq3/
    qualitative/

Every chart is saved as:
    - PDF
    - PNG (300 DPI)

Design principles
-----------------
* Pie charts are used for mutually exclusive part-to-whole distributions.
* Multiple-selection questions use horizontal bar charts because their
  percentages may sum to more than 100%.
* Roles are an exception: both a participant-percentage bar chart and a
  composition-of-role-selections pie chart are generated.
* Likert questions use 100% stacked horizontal bars.
* Percentages are placed inside bars/slices when space permits, using
  contrasting text colors.
"""

from __future__ import annotations

import argparse
import re
import textwrap
from collections import Counter
from pathlib import Path
from typing import Dict, Mapping, Sequence

import matplotlib.pyplot as plt
import numpy as np

from survey_statistics import (
    PRACTICES,
    DEFAULT_INITIAL_CANDIDATES,
    DEFAULT_AFTER_BOT_FILTER,
    DEFAULT_CONTACTED,
    compute_all,
    read_input,
)


# =============================================================================
# Colors
# =============================================================================

BLUE = "#6FA8DC"
DARK_BLUE = "#4F81BD"
LIGHT_BLUE = "#A9C9E8"

GRAY = "#B8B8B8"
LIGHT_GRAY = "#E1E1E1"
DARK_GRAY = "#777777"

ORANGE = "#F6A04D"

RED = "#D96B6B"
DARK_RED = "#C0504D"

GREEN = "#79A96B"
PURPLE = "#8E7CC3"
TEAL = "#76A5AF"


PIE_COLORS = [
    BLUE,
    ORANGE,
    GREEN,
    RED,
    PURPLE,
    TEAL,
    GRAY,
    LIGHT_BLUE,
]


LIKERT_ORDER = [
    "Strongly disagree",
    "Disagree",
    "Neutral",
    "Agree",
    "Strongly agree",
]


LIKERT_COLORS = {
    "Strongly disagree": DARK_RED,
    "Disagree": ORANGE,
    "Neutral": GRAY,
    "Agree": BLUE,
    "Strongly agree": DARK_BLUE,
}


# =============================================================================
# Display labels
# =============================================================================

PRACTICE_LABELS = {
    "G1": "G1",
    "G2": "G2",
    "G3": "G3",
    "G4": "G4",
    "G5": "G5",
    "G6": "G6",
    "G7": "G7",

    "P1": "P1",
    "P2": "P2",
    "P3": "P3",
    "P4": "P4",
    "P5": "P5",
}


SHORT_LABELS = {
    # -------------------------------------------------------------------------
    # Roles
    # -------------------------------------------------------------------------
    "Organization/team member involved in template design":
        "Organization/team member",

    "User who later contributed improvements":
        "User who later contributed",

    # -------------------------------------------------------------------------
    # Motivations
    # -------------------------------------------------------------------------
    "Reusing project structure across multiple projects":
        "Reuse project structure",

    "Reducing setup/configuration effort for others":
        "Reduce setup/configuration effort",

    "Standardizing development practices":
        "Standardize development practices",

    "Demonstrating a reusable project structure":
        "Demonstrate reusable structure",

    "Sharing best practices with the community":
        "Share best practices",

    "Encouraging consistent tooling, such as testing, linting, or CI/CD":
        "Encourage consistent tooling",

    "Supporting a specific technology, framework, or ecosystem":
        "Support a technology/framework",

    "Supporting team onboarding":
        "Support team onboarding",

    "Supporting teaching or training":
        "Support teaching/training",

    # -------------------------------------------------------------------------
    # Contexts
    # -------------------------------------------------------------------------
    "Industry/professional projects":
        "Industry/professional projects",

    "Internal organizational reuse":
        "Internal organizational reuse",

    "Experimentation or prototyping":
        "Experimentation/prototyping",

    "Teaching or educational activities":
        "Teaching/education",

    # -------------------------------------------------------------------------
    # Audiences
    # -------------------------------------------------------------------------
    "External developers/community":
        "External developers/community",

    "Open-source contributors":
        "Open-source contributors",

    "My organization/company":
        "My organization/company",

    # -------------------------------------------------------------------------
    # Maintenance activities
    # -------------------------------------------------------------------------
    "Keeping dependencies up to date":
        "Keep dependencies up to date",

    "Fixing reported issues":
        "Fix reported issues",

    "Updating documentation":
        "Update documentation",

    "Monitoring vulnerabilities or security alerts":
        "Monitor vulnerabilities/security",

    "Updating supported technology versions":
        "Update supported technology versions",

    "Maintaining CI/CD pipelines":
        "Maintain CI/CD pipelines",

    "Responding to user questions":
        "Respond to user questions",

    "Publishing releases or changelogs":
        "Publish releases/changelogs",

    "Maintaining tests":
        "Maintain tests",

    # -------------------------------------------------------------------------
    # Maintenance signals
    # -------------------------------------------------------------------------
    "Recent commits":
        "Recent commits",

    "Dependency updates":
        "Dependency updates",

    "Issue responses":
        "Issue responses",

    "Security fixes":
        "Security fixes",

    "Recent releases":
        "Recent releases",

    "Pull request reviews":
        "Pull request reviews",

    "Technology/version updates":
        "Technology/version updates",

    "Documentation updates":
        "Documentation updates",

    "CI/CD updates":
        "CI/CD updates",

    # -------------------------------------------------------------------------
    # Quality monitoring
    # -------------------------------------------------------------------------
    "Security alerts/dependency risks":
        "Security alerts/dependency risks",

    "Code smells/maintainability issues":
        "Code smells/maintainability",

    "Outdated dependencies":
        "Outdated dependencies",

    "Build/CI failures":
        "Build/CI failures",

    "Do not actively monitor":
        "Do not actively monitor",

    # -------------------------------------------------------------------------
    # P1 adaptations
    # -------------------------------------------------------------------------
    "Rewrote or updated README":
        "Rewrite/update README",

    "Added setup/installation/configuration instructions":
        "Add setup/configuration instructions",

    "Added supporting files":
        "Add supporting files",

    "Removed project-specific content":
        "Remove project-specific content",

    "Replaced real configuration files with examples":
        "Replace configurations with examples",

    "Removed sensitive information":
        "Remove sensitive information",

    "Added examples/comments/guidance":
        "Add examples/guidance",

    "Added or updated automation":
        "Add/update automation",

    "Replaced fixed names/paths/URLs/project information with placeholders":
        "Replace fixed values with placeholders",

    "No specific changes":
        "No specific changes",

    # -------------------------------------------------------------------------
    # AI
    # -------------------------------------------------------------------------
    "Generate/improve documentation":
        "Generate/improve documentation",

    "Review code/detect issues/suggest improvements":
        "Review code/detect issues",

    "Improve/refactor template structure":
        "Improve/refactor structure",

    "Configure CI/CD/tests/linters/dependency management":
        "Configure CI/CD/tests/dependencies",

    "Generate initial template structure":
        "Generate initial structure",

    "Create templates for AI-based applications":
        "Create AI-application templates",

    "No AI use":
        "No AI use",

    "Not sure":
        "Not sure",

    "Regular Project":
        "Non-specific Domain",
}


QUALITATIVE_LABELS = {
    "NoProblem":
        "No problem reported",

    "KeepUpToDateGeneratedRepos":
        "Keep generated repositories up to date",

    "AutomationSetup":
        "Automation/setup",

    "BuildComunity":
        "Building community",

    "CustomizationEffort":
        "Customization effort",

    "UpdateMigrateDependencyTech":
        "Dependency/technology migration",

    "EvolutionOfTemplates":
        "Template evolution",

    "Documentation":
        "Documentation",

    "CreatingWhileUpdatingTemplate":
        "Configuration during project creation",

    "BetterDocumentation":
        "Better documentation",

    "ClearUpgradeStrategy":
        "Clear upgrade strategy",

    "TrackTemplateAndTheirRepos":
        "Track template/generated repo relation",

    "BetterSearch":
        "Better template search",

    "TemplateVersioning":
        "Template versioning",

    "CommunicationNotifyUsersAboutChanges":
        "Notify users about changes",
}

SHORT_LABELS.update({
    "Software Engineer":
        "Software Engineer",

    "Senior Software Engineer":
        "Senior Software Engineer",

    "Software Developer":
        "Software Developer",

    "Researcher":
        "Researcher",

    "Professor":
        "Professor",

    "Graduate Student":
        "Graduate Student",

    "Undergraduate Student":
        "Undergraduate Student",

    "Engineering Manager":
        "Engineering Manager",
})


# =============================================================================
# Generic utilities
# =============================================================================

def percentage(part, total):
    return part / total * 100 if total else 0.0


def short_label(value):
    return SHORT_LABELS.get(
        value,
        value,
    )


def wrap_label(
    value,
    width=32,
):
    return "\n".join(
        textwrap.wrap(
            str(value),
            width=width,
            break_long_words=False,
            break_on_hyphens=False,
        )
    )


def humanize_code(code):
    if code in QUALITATIVE_LABELS:
        return QUALITATIVE_LABELS[code]

    text = re.sub(
        r"([a-z0-9])([A-Z])",
        r"\1 \2",
        code,
    )

    return text.replace(
        "_",
        " ",
    )


def contrast_text_color(hex_color):
    """
    Return white or black depending on perceived brightness.
    """

    color = hex_color.lstrip("#")

    r = int(
        color[0:2],
        16,
    )

    g = int(
        color[2:4],
        16,
    )

    b = int(
        color[4:6],
        16,
    )

    luminance = (
        0.299 * r
        + 0.587 * g
        + 0.114 * b
    )

    if luminance > 165:
        return "black"

    return "white"


def rgb_contrast_text_color(facecolor):
    """
    Same idea as contrast_text_color, but receives RGBA from Matplotlib.
    """

    r, g, b, _ = facecolor

    luminance = (
        0.299 * r
        + 0.587 * g
        + 0.114 * b
    )

    if luminance > 0.65:
        return "black"

    return "white"


def add_title(
    ax,
    title,
):
    if title:
        ax.set_title(
            title,
            fontsize=14,
            pad=14,
        )


def clean_axes(
    ax,
    axis,
):
    ax.grid(
        axis=axis,
        alpha=0.20,
        linewidth=0.8,
    )

    ax.set_axisbelow(
        True
    )

    ax.spines["top"].set_visible(
        False
    )

    ax.spines["right"].set_visible(
        False
    )


def save_figure(
    fig,
    output_dir,
    name,
    *,
    tight=True,
):
    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    png = output_dir / f"{name}.png"
    pdf = output_dir / f"{name}.pdf"

    save_kwargs = {}

    if tight:
        save_kwargs["bbox_inches"] = "tight"

    fig.savefig(
        png,
        dpi=300,
        **save_kwargs,
    )

    fig.savefig(
        pdf,
        **save_kwargs,
    )

    plt.close(fig)

    return [
        png,
        pdf,
    ]


def merge_distribution_categories(
    counts,
    groups,
):
    """
    Merge alternate spellings such as:
        5-10 years
        5–10 years

    groups:
        {
            "5–10 years": ["5–10 years", "5-10 years"]
        }
    """

    result = {}

    used = set()

    for target, alternatives in groups.items():

        result[target] = sum(
            counts.get(
                alternative,
                0,
            )
            for alternative in alternatives
        )

        used.update(
            alternatives
        )

    for key, value in counts.items():

        if (
            key not in used
            and key != "<blank>"
            and value > 0
        ):
            result[key] = (
                result.get(
                    key,
                    0,
                )
                + value
            )

    return result


# =============================================================================
# Horizontal bar charts
# =============================================================================

def generate_horizontal_bar(
    counts,
    denominator,
    output_dir,
    name,
    title="",
    *,
    color=BLUE,
    max_items=None,
    label_width=34,
):
    """
    Horizontal percentage bar chart.

    Appropriate for multiple-selection questions, where totals may exceed 100%.

    All values are displayed inside the bars. Font size is reduced slightly
    for shorter bars to preserve readability.
    """

    data = [
        (item, count)
        for item, count
        in counts.items()
        if count > 0
    ]

    data.sort(
        key=lambda item: item[1],
        reverse=True,
    )

    if max_items is not None:
        data = data[:max_items]

    # Reverse so the largest category appears at the top.
    data = data[::-1]

    labels = [
        wrap_label(
            short_label(item),
            label_width,
        )
        for item, _
        in data
    ]

    raw_counts = np.array(
        [
            count
            for _, count
            in data
        ],
        dtype=float,
    )

    percentages = (
        raw_counts
        / denominator
        * 100
    )

    height = max(
        4.0,
        len(data) * 0.50 + 1.35,
    )

    fig, ax = plt.subplots(
        figsize=(9.0, height)
    )

    y = np.arange(
        len(data)
    )

    bars = ax.barh(
        y,
        percentages,
        height=0.68,
        color=color,
        edgecolor="white",
        linewidth=0.7,
    )

    ax.set_yticks(y)

    ax.set_yticklabels(
        labels,
        fontsize=11,
    )

    max_value = (
        float(max(percentages))
        if len(percentages)
        else 0
    )

    # Keep the axis compact while retaining a percentage reference.
    upper = min(
        105,
        max_value + 8,
    )

    ax.set_xlim(
        0,
        upper,
    )

    ticks = [
        value
        for value
        in [0, 20, 40, 60, 80, 100]
        if value <= upper
    ]

    ax.set_xticks(
        ticks
    )

    ax.set_xticklabels(
        [
            f"{value}%"
            for value
            in ticks
        ]
    )

    ax.set_xlabel(
        "Participants (%)",
        fontsize=10,
    )

    clean_axes(
        ax,
        axis="x",
    )

    inside_color = contrast_text_color(
        color
    )

    for bar, count, value in zip(
            bars,
            raw_counts,
            percentages,
    ):

        # Large bars: count + percentage inside.
        if value >= 18:

            label = (
                f"{int(count)} "
                f"({value:.1f}%)"
            )

            ax.text(
                value - 1.0,
                bar.get_y()
                + bar.get_height() / 2,
                label,
                ha="right",
                va="center",
                fontsize=10,
                color=inside_color,
                fontweight="bold",
            )

        # Medium/small bars: percentage inside.
        elif value >= 4:

            label = f"{value:.1f}%"

            ax.text(
                value - 0.8,
                bar.get_y()
                + bar.get_height() / 2,
                label,
                ha="right",
                va="center",
                fontsize=9,
                color=inside_color,
                fontweight="bold",
            )

        # Very small bars: percentage outside in black.
        else:

            label = f"{value:.1f}%"

            ax.text(
                value + 0.6,
                bar.get_y()
                + bar.get_height() / 2,
                label,
                ha="left",
                va="center",
                fontsize=7.5,
                color="black",
                fontweight="bold",
            )

    add_title(
        ax,
        title,
    )

    fig.subplots_adjust(
        left=0.39,
        right=0.97,
        bottom=0.12,
        top=0.88,
    )

    return save_figure(
        fig,
        output_dir,
        name,
    )

# =============================================================================
# Vertical bar charts
# =============================================================================

def generate_vertical_bar(
    counts,
    denominator,
    order,
    output_dir,
    name,
    title="",
    *,
    label_map=None,
    color=BLUE,
):
    """
    Vertical percentage bars for ordinal or compact categorical distributions.
    """

    rows = []

    used = set()

    for category in order:

        if category in used:
            continue

        value = counts.get(
            category,
            0,
        )

        if value > 0:

            rows.append(
                (
                    category,
                    value,
                )
            )

            used.add(
                category
            )

    # Add any unexpected non-empty categories.
    for category, value in counts.items():

        if (
            category not in used
            and category != "<blank>"
            and value > 0
        ):

            rows.append(
                (
                    category,
                    value,
                )
            )

    labels = []

    raw_counts = []

    for category, count in rows:

        if label_map:

            label = label_map.get(
                category,
                short_label(category),
            )

        else:

            label = short_label(
                category
            )

        labels.append(
            wrap_label(
                label,
                18,
            )
        )

        raw_counts.append(
            count
        )

    raw_counts = np.asarray(
        raw_counts,
        dtype=float,
    )

    percentages = (
        raw_counts
        / denominator
        * 100
    )

    x = np.arange(
        len(labels)
    )

    fig, ax = plt.subplots(
        figsize=(
            8.4,
            5.5,
        )
    )

    bars = ax.bar(
        x,
        percentages,
        width=0.68,
        color=color,
        edgecolor="white",
        linewidth=0.7,
    )

    ax.set_xticks(
        x
    )

    ax.set_xticklabels(
        labels,
        fontsize=9,
    )

    max_value = (
        float(
            max(percentages)
        )
        if len(percentages)
        else 0
    )

    upper = min(
        105,
        max_value + 14,
    )

    ax.set_ylim(
        0,
        upper,
    )

    ticks = [
        value
        for value
        in [
            0,
            20,
            40,
            60,
            80,
            100,
        ]
        if value <= upper
    ]

    ax.set_yticks(
        ticks
    )

    ax.set_yticklabels(
        [
            f"{value}%"
            for value
            in ticks
        ]
    )

    ax.set_ylabel(
        "Participants (%)",
        fontsize=10,
    )

    clean_axes(
        ax,
        axis="y",
    )

    inside_color = (
        contrast_text_color(
            color
        )
    )

    for bar, count, value in zip(
        bars,
        raw_counts,
        percentages,
    ):

        if value >= 17:

            ax.text(
                bar.get_x()
                + bar.get_width() / 2,
                value - 2,
                (
                    f"{int(count)}\n"
                    f"{value:.1f}%"
                ),
                ha="center",
                va="top",
                fontsize=8.5,
                color=inside_color,
                fontweight="bold",
            )

        else:

            ax.text(
                bar.get_x()
                + bar.get_width() / 2,
                value + 1,
                (
                    f"{int(count)}\n"
                    f"{value:.1f}%"
                ),
                ha="center",
                va="bottom",
                fontsize=8.5,
                color="black",
            )

    add_title(
        ax,
        title,
    )

    fig.subplots_adjust(
        left=0.10,
        right=0.98,
        bottom=0.20,
        top=0.87,
    )

    return save_figure(
        fig,
        output_dir,
        name,
    )


# =============================================================================
# Pie charts
# =============================================================================

def generate_pie(
    counts,
    order,
    output_dir,
    name,
    title="",
    *,
    label_map=None,
    center_text=None,
    denominator=None,
    denominator_name="",
):
    """
    Pie/donut chart for true part-to-whole distributions.

    Percentages are based on the sum of displayed counts unless an explicit
    denominator is provided.
    """

    rows = []

    for category in order:

        count = counts.get(
            category,
            0,
        )

        if count > 0:

            rows.append(
                (
                    category,
                    count,
                )
            )

    labels = []

    values = []

    for category, count in rows:

        if label_map:

            label = label_map.get(
                category,
                short_label(category),
            )

        else:

            label = short_label(
                category
            )

        labels.append(
            label
        )

        values.append(
            count
        )

    if not values:
        return []

    total = (
        denominator
        if denominator is not None
        else sum(values)
    )

    colors = PIE_COLORS[
        :len(values)
    ]

    fig, ax = plt.subplots(
        figsize=(
            8.2,
            5.8,
        )
    )

    wedges, _, autotexts = ax.pie(
        values,
        colors=colors,
        startangle=90,
        counterclock=False,

        autopct=lambda p:
            f"{p:.1f}%"
            if p >= 4
            else "",

        pctdistance=0.72,

        wedgeprops={
            "width": 0.58,
            "edgecolor": "white",
            "linewidth": 1.2,
        },

        textprops={
            "fontsize": 9,
        },
    )

    for wedge, text in zip(
        wedges,
        autotexts,
    ):

        text.set_color(
            rgb_contrast_text_color(
                wedge.get_facecolor()
            )
        )

        text.set_fontweight(
            "bold"
        )

        text.set_fontsize(
            12
        )

    if center_text is None:

        center_text = (
            f"N={total}"
        )

    ax.text(
        0,
        0.04,
        center_text,
        ha="center",
        va="center",
        fontsize=12,
        fontweight="bold",
    )

    if "\n" not in center_text:

        ax.text(
            0,
            -0.12,
            denominator_name,
            ha="center",
            va="center",
            fontsize=8,
        )

    legend_labels = [
        wrap_label(label, 24)
        for label in labels
    ]

    legend = fig.legend(
        wedges,
        legend_labels,
        loc="center right",
        bbox_to_anchor=(
            0.985,
            0.5,
        ),
        frameon=True,
        fontsize=11,
        fancybox=False,
        framealpha=1.0,
        labelspacing=0.3,
    )

    legend.get_frame().set_linewidth(
        0.8
    )

    legend.get_frame().set_edgecolor(
        "0.55"
    )

    add_title(
        ax,
        title,
    )

    ax.axis(
        "equal"
    )

    fig.subplots_adjust(
        left=0.03,
        right=0.71,
        bottom=0.06,
        top=0.87,
    )

    return save_figure(
        fig,
        output_dir,
        name,
        tight=False,
    )


# =============================================================================
# Multi-select composition pie
# =============================================================================

def generate_multiselect_composition_pie(
    counts,
    output_dir,
    name,
    title="",
):
    """
    Pie chart for a multiple-selection question.

    IMPORTANT:
    percentages represent the composition of all reported selections,
    NOT the percentage of respondents.

    This is useful for roles, for example, as long as the figure caption
    explicitly explains the denominator.
    """

    cleaned = {
        key: value
        for key, value
        in counts.items()
        if value > 0
    }

    ordered = sorted(
        cleaned,
        key=lambda key:
            cleaned[key],
        reverse=True,
    )

    total_selections = sum(
        cleaned.values()
    )

    return generate_pie(
        cleaned,
        ordered,
        output_dir,
        name,
        title,
        center_text=(
            f"{total_selections}\n"
            "selections"
        ),
        denominator=total_selections,
        denominator_name="selections",
    )


# =============================================================================
# Likert charts
# =============================================================================

def generate_likert(
    stats,
    practices,
    output_dir,
    name,
    title="",
):
    """
    100% stacked horizontal Likert chart.

    Uses the same visual logic as the stacked evolution figures:
      - fixed percentage axis
      - white separators
      - labels inside sufficiently large segments
      - restrained legend on the right
    """

    n = stats[
        "data_collection"
    ][
        "final_analytical_sample"
    ]

    labels = [
        PRACTICE_LABELS[
            practice
        ]
        for practice
        in practices
    ]

    y = np.arange(
        len(practices)
    )

    values_by_category = {}

    for category in LIKERT_ORDER:

        values_by_category[
            category
        ] = np.array(
            [
                percentage(
                    stats[
                        "likert"
                    ][
                        practice
                    ][
                        "counts"
                    ][
                        category
                    ],
                    n,
                )
                for practice
                in practices
            ],
            dtype=float,
        )

    height = max(
        4.5,
        0.52 * len(practices)
        + 1.4,
    )

    fig, ax = plt.subplots(
        figsize=(
            9.6,
            height,
        )
    )

    left = np.zeros(
        len(practices),
        dtype=float,
    )

    for category in LIKERT_ORDER:

        values = values_by_category[
            category
        ]

        bars = ax.barh(
            y,
            values,
            height=0.68,
            left=left,
            label=category,
            color=LIKERT_COLORS[
                category
            ],
            edgecolor="white",
            linewidth=0.7,
        )

        text_color = (
            contrast_text_color(
                LIKERT_COLORS[
                    category
                ]
            )
        )

        for i, value in enumerate(
            values
        ):

            if value >= 7:

                ax.text(
                    left[i]
                    + value / 2,
                    y[i],
                    f"{value:.1f}%",
                    ha="center",
                    va="center",
                    fontsize=8,
                    color=text_color,
                    fontweight="bold",
                )

        left += values

    ax.set_xlim(
        0,
        100,
    )

    ax.set_xticks(
        [
            0,
            25,
            50,
            75,
            100,
        ]
    )

    ax.set_xticklabels(
        [
            "0%",
            "25%",
            "50%",
            "75%",
            "100%",
        ]
    )

    ax.set_xlabel(
        "Participants (%)",
        fontsize=10,
    )

    ax.set_yticks(
        y
    )

    ax.set_yticklabels(
        labels,
        fontsize=10,
        fontweight="bold",
    )

    ax.invert_yaxis()

    clean_axes(
        ax,
        axis="x",
    )

    add_title(
        ax,
        title,
    )

    legend = ax.legend(
        loc="center left",
        bbox_to_anchor=(
            1.015,
            0.5,
        ),
        frameon=True,
        title="Response",
        fontsize=9,
        title_fontsize=10,
        fancybox=False,
        framealpha=1.0,
    )

    legend.get_frame().set_linewidth(
        0.8
    )

    legend.get_frame().set_edgecolor(
        "0.55"
    )

    fig.subplots_adjust(
        left=0.09,
        right=0.75,
        bottom=0.13,
        top=0.86,
    )

    return save_figure(
        fig,
        output_dir,
        name,
    )


# =============================================================================
# Qualitative codes
# =============================================================================

def generate_qualitative_chart(
    stats,
    question,
    output_dir,
    top_k=10,
):
    info = stats[
        "qualitative_codes"
    ].get(
        question
    )

    if not info:
        return []

    denominator = info[
        "responses_with_codes"
    ]

    if denominator == 0:
        return []

    counts = Counter(
        info[
            "counts"
        ]
    )

    pretty_counts = {
        humanize_code(code):
            count
        for code, count
        in counts.items()
    }

    return generate_horizontal_bar(
        pretty_counts,
        denominator,
        output_dir,
        f"qualitative_{question.lower()}",
        f"{question}: Most Frequent Qualitative Themes",
        max_items=top_k,
        label_width=34,
        color=BLUE,
    )


# =============================================================================
# Generate all figures
# =============================================================================

def generate_all(
    stats,
    output_dir,
    top_codes=10,
):
    n = stats[
        "data_collection"
    ][
        "final_analytical_sample"
    ]

    figure_paths = []

    dc_dir = (
        output_dir
        / "data_collection"
    )

    rq1_dir = (
        output_dir
        / "rq1"
    )

    rq2_dir = (
        output_dir
        / "rq2"
    )

    rq3_dir = (
        output_dir
        / "rq3"
    )

    qualitative_dir = (
        output_dir
        / "qualitative"
    )

    # =========================================================================
    # DATA COLLECTION
    # =========================================================================

    # -------------------------------------------------------------------------
    # Roles
    #
    # Generate both:
    # 1. participant prevalence
    # 2. composition of role selections
    # -------------------------------------------------------------------------

    role_order = [
        "Creator",
        "Maintainer",
        "Contributor",
        "Organization/team member involved in template design",
        "User who later contributed improvements",
    ]

    figure_paths += generate_vertical_bar(
        stats[
            "multi_select"
        ][
            "roles"
        ],
        n,
        role_order,
        dc_dir,
        "participant_roles",
        #"Participants' Roles in Template Repositories",
        color=BLUE,
    )

    figure_paths += generate_multiselect_composition_pie(
        stats[
            "multi_select"
        ][
            "roles"
        ],
        dc_dir,
        "participant_roles_composition",
        #"Composition of Reported Template Repository Roles",
    )

    # -------------------------------------------------------------------------
    # GitHub usage frequency
    # -------------------------------------------------------------------------

    github_order = [
        "Daily",
        "Frequently",
        "Occasionally",
        "Rarely",
    ]

    figure_paths += generate_pie(
        stats[
            "distributions"
        ][
            "github_frequency"
        ],
        github_order,
        dc_dir,
        "github_usage_frequency",
        #"GitHub Usage Frequency",
        center_text=f"N={n}",
    )

    # -------------------------------------------------------------------------
    # Software development experience
    # -------------------------------------------------------------------------

    raw_development = stats[
        "distributions"
    ][
        "years_dev"
    ]

    development_experience = (
        merge_distribution_categories(
            raw_development,
            {
                "Less than 1 year": [
                    "Less than 1 year",
                ],

                "1–3 years": [
                    "1–3 years",
                    "1-3 years",
                ],

                "3–5 years": [
                    "3–5 years",
                    "3-5 years",
                ],

                "5–10 years": [
                    "5–10 years",
                    "5-10 years",
                ],

                ">10 years": [
                    "More than 10 years",
                ],
            },
        )
    )

    figure_paths += generate_pie(
        development_experience,
        [
            "Less than 1 year",
            "1–3 years",
            "3–5 years",
            "5–10 years",
            ">10 years",
        ],
        dc_dir,
        "software_development_experience",
        #"Software Development Experience",
        center_text=f"N={n}",
    )

    # -------------------------------------------------------------------------
    # Number of templates
    # -------------------------------------------------------------------------

    raw_templates = stats[
        "distributions"
    ][
        "num_templates"
    ]

    template_experience = (
        merge_distribution_categories(
            raw_templates,
            {
                "1": [
                    "1",
                ],

                "2–5": [
                    "2–5",
                    "2-5",
                ],

                "5–10": [
                    "5–10",
                    "5-10",
                ],

                "10–20": [
                    "10–20",
                    "10-20",
                ],

                ">20": [
                    "More than 20",
                ],
            },
        )
    )

    figure_paths += generate_pie(
        template_experience,
        [
            "1",
            "2–5",
            "5–10",
            "10–20",
            ">20",
        ],
        dc_dir,
        "template_repository_experience",
        #"Experience with GitHub Template Repositories",
        center_text=f"N={n}",
    )

    # -------------------------------------------------------------------------
    # Current positions
    # -------------------------------------------------------------------------

    current_positions = {
        key: value
        for key, value
        in stats[
            "distributions"
        ][
            "current_position"
        ].items()
        if (
                key != "<blank>"
                and value > 0
        )
    }

    # Group Unemployed and Hobbies as Others
    others = (
            current_positions.pop("Unemployed", 0)
            + current_positions.pop("Hobbyist", 0)
    )

    if others > 0:
        current_positions["Others"] = (
                current_positions.get("Others", 0)
                + others
        )

    current_position_order = [
        key
        for key, _
        in sorted(
            current_positions.items(),
            key=lambda item: item[1],
            reverse=True,
        )
    ]

    figure_paths += generate_pie(
        current_positions,
        current_position_order,
        dc_dir,
        "current_positions",
        #"Participants' Current Positions",
        center_text=f"N={n}",
    )

    figure_paths += generate_pie(
        current_positions,
        current_position_order,
        dc_dir,
        "participant_positions",
        #"Participants' Positions",
        center_text=f"N={n}",
    )

    # =========================================================================
    # RQ1
    # =========================================================================

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "motivations"
        ],
        n,
        rq1_dir,
        "motivations",
        "Motivations for Template Repository Involvement",
        color=BLUE,
        label_width=34,
    )

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "contexts"
        ],
        n,
        rq1_dir,
        "contexts",
        "Contexts of Template Repository Involvement",
        color=BLUE,
        label_width=34,
    )

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "domains"
        ],
        n,
        rq1_dir,
        "practitioner_reported_domains",
        "Practitioner-Reported Template Domains",
        color=BLUE,
        label_width=34,
    )

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "audiences"
        ],
        n,
        rq1_dir,
        "intended_audiences",
        "Intended Audiences of Template Repositories",
        color=BLUE,
        label_width=34,
    )

    # =========================================================================
    # RQ2
    # =========================================================================

    # -------------------------------------------------------------------------
    # Maintenance frequency
    # -------------------------------------------------------------------------

    maintenance_labels = {
        "I update them regularly":
            "Regularly",

        "I update them occasionally":
            "Occasionally",

        "I update them only when needed":
            "Only when needed",

        "I no longer maintain them":
            "No longer maintain",

        "Maintenance is shared with others":
            "Shared maintenance",

        "I am not sure":
            "Not sure",
    }

    maintenance_counts = {}

    for key, value in stats[
        "distributions"
    ][
        "maintenance_active"
    ].items():

        if (
            key != "<blank>"
            and value > 0
        ):

            maintenance_counts[
                maintenance_labels.get(
                    key,
                    key,
                )
            ] = value

    maintenance_order = [
        "Regularly",
        "Occasionally",
        "Only when needed",
        "No longer maintain",
        "Shared maintenance",
        "Not sure",
    ]

    figure_paths += generate_pie(
        maintenance_counts,
        maintenance_order,
        rq2_dir,
        "maintenance_frequency",
        "How Actively Are Template Repositories Maintained?",
        center_text=f"N={n}",
    )

    # -------------------------------------------------------------------------
    # Maintenance activities
    # -------------------------------------------------------------------------

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "maintenance_aspects"
        ],
        n,
        rq2_dir,
        "important_maintenance_activities",
        "Important Template Maintenance Activities",
        color=BLUE,
        label_width=34,
    )

    # -------------------------------------------------------------------------
    # Maintenance signals
    # -------------------------------------------------------------------------

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "maintenance_signals"
        ],
        n,
        rq2_dir,
        "maintenance_signals",
        "Meaningful Evidence of Template Maintenance",
        color=BLUE,
        label_width=34,
    )

    # -------------------------------------------------------------------------
    # Ownership perceptions
    # -------------------------------------------------------------------------

    ownership_raw = stats[
        "distributions"
    ][
        "ownership"
    ]

    ownership_counts = {
        "Organizations usually differ":
            ownership_raw.get(
                "Yes, templates maintained by organizations usually follow different maintenance practices",
                0,
            ),

        "Individuals usually differ":
            ownership_raw.get(
                "Yes, templates maintained by individuals usually follow different maintenance practices",
                0,
            ),

        "Ownership does not matter":
            ownership_raw.get(
                "No, ownership does not influence maintenance practices",
                0,
            ),

        "It depends":
            ownership_raw.get(
                "It depends",
                0,
            ),

        "Not sure":
            ownership_raw.get(
                "I am not sure",
                0,
            ),
    }

    figure_paths += generate_pie(
        ownership_counts,
        [
            "Organizations usually differ",
            "Individuals usually differ",
            "Ownership does not matter",
            "It depends",
            "Not sure",
        ],
        rq2_dir,
        "ownership_and_maintenance",
        #"Perceived Influence of Ownership on Maintenance",
        center_text=f"N={n}",
    )

    # -------------------------------------------------------------------------
    # Quality propagation
    # -------------------------------------------------------------------------

    propagation_raw = stats[
        "distributions"
    ][
        "quality_propagation"
    ]

    propagation_counts = {
        "Yes":
            propagation_raw.get(
                "Yes",
                0,
            ),

        "It depends":
            propagation_raw.get(
                "It depends",
                0,
            ),

        "No":
            propagation_raw.get(
                "No",
                0,
            ),

        "Not sure":
            propagation_raw.get(
                "I am not sure",
                0,
            ),
    }

    figure_paths += generate_pie(
        propagation_counts,
        [
            "Yes",
            "It depends",
            "No",
            "Not sure",
        ],
        rq2_dir,
        "quality_issue_propagation",
        #"Can Template Quality Issues Affect Generated Projects?",
        center_text=f"N={n}",
    )

    # -------------------------------------------------------------------------
    # Quality monitoring
    # -------------------------------------------------------------------------

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "quality_monitoring"
        ],
        n,
        rq2_dir,
        "quality_monitoring",
        "Quality Issues Actively Monitored",
        color=BLUE,
        label_width=34,
    )

    # =========================================================================
    # RQ3
    # =========================================================================

    # -------------------------------------------------------------------------
    # Guidelines
    # -------------------------------------------------------------------------

    figure_paths += generate_likert(
        stats,
        [
            "G1",
            "G2",
            "G3",
            "G4",
            "G5",
            "G6",
            "G7",
        ],
        rq3_dir,
        "guidelines_likert",
        "Practitioner Assessment of the Proposed Guidelines",
    )

    # -------------------------------------------------------------------------
    # Pitfalls
    # -------------------------------------------------------------------------

    figure_paths += generate_likert(
        stats,
        [
            "P1",
            "P2",
            "P3",
            "P4",
            "P5",
        ],
        rq3_dir,
        "pitfalls_likert",
        "Practitioner Assessment of the Identified Pitfalls",
    )

    # -------------------------------------------------------------------------
    # Combined Likert
    # -------------------------------------------------------------------------

    figure_paths += generate_likert(
        stats,
        PRACTICES,
        rq3_dir,
        "guidelines_and_pitfalls_likert",
        "Practitioner Assessment of Guidelines and Pitfalls",
    )

    # -------------------------------------------------------------------------
    # P1 conversion adaptations
    # -------------------------------------------------------------------------

    p1_n = stats[
        "derived"
    ][
        "p1_conversion_subset_n"
    ]

    p1_counts = {
        key: value
        for key, value
        in stats[
            "derived"
        ][
            "p1_conversion_adaptations"
        ].items()
        if (
            key
            != "Created as a template from the beginning"
            and value > 0
        )
    }

    figure_paths += generate_horizontal_bar(
        p1_counts,
        p1_n,
        rq3_dir,
        "repository_to_template_adaptations",
        "Changes Made When Converting Repositories into Templates",
        color=BLUE,
        label_width=35,
    )

    # -------------------------------------------------------------------------
    # AI use
    # -------------------------------------------------------------------------

    figure_paths += generate_horizontal_bar(
        stats[
            "multi_select"
        ][
            "ai"
        ],
        n,
        rq3_dir,
        "ai_tool_usage",
        "AI Tool Use in Template Repository Work",
        color=BLUE,
        label_width=34,
    )

    # =========================================================================
    # Qualitative analysis
    # =========================================================================

    for question in [
        "OQ1",
        "OQ2",
        "OQ3",
        "OQ4",
        "OQ5",
        "OQ6",
    ]:

        figure_paths += (
            generate_qualitative_chart(
                stats,
                question,
                qualitative_dir,
                top_k=top_codes,
            )
        )

    return figure_paths


# =============================================================================
# Main
# =============================================================================

def main():

    parser = argparse.ArgumentParser(
        description=(
            "Generate publication-quality figures for the "
            "GitHub Template Repositories practitioner survey."
        )
    )

    parser.add_argument(
        "input",
        type=Path,
        help=(
            "Survey response workbook "
            "(.xlsx or .csv)."
        ),
    )

    parser.add_argument(
        "--sheet",
        default=None,
        help=(
            "Optional worksheet name. "
            "Defaults to the first worksheet."
        ),
    )

    parser.add_argument(
        "--output_dir",
        type=Path,
        default=Path(
            "survey_charts"
        ),
        help=(
            "Output directory "
            "(default: survey_charts)."
        ),
    )

    parser.add_argument(
        "--top_codes",
        type=int,
        default=10,
        help=(
            "Maximum number of qualitative "
            "codes shown per OQ chart."
        ),
    )

    parser.add_argument(
        "--initial_candidates",
        type=int,
        default=DEFAULT_INITIAL_CANDIDATES,
    )

    parser.add_argument(
        "--after_bot_filter",
        type=int,
        default=DEFAULT_AFTER_BOT_FILTER,
    )

    parser.add_argument(
        "--contacted",
        type=int,
        default=DEFAULT_CONTACTED,
    )

    parser.add_argument(
        "--undeliverable",
        type=int,
        default=None,
    )

    args = parser.parse_args()

    headers, rows = read_input(
        args.input,
        sheet_name=args.sheet,
    )

    stats = compute_all(
        headers,
        rows,
        initial_candidates=args.initial_candidates,
        after_bot_filter=args.after_bot_filter,
        contacted=args.contacted,
        undeliverable=args.undeliverable,
    )

    print()
    print(
        "Survey figures"
    )
    print(
        "-" * 80
    )

    print(
        "Analytical sample: "
        f"N={stats['data_collection']['final_analytical_sample']}"
    )

    figure_paths = generate_all(
        stats,
        args.output_dir,
        top_codes=args.top_codes,
    )

    print()
    print(
        "Saved figures:"
    )

    for path in figure_paths:
        print(
            f"- {path}"
        )


if __name__ == "__main__":
    main()