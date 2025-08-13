#!/usr/bin/env python3
"""Main CLI for DOCX Navigator."""

import sys
from pathlib import Path
from typing import List

import click

from .merger import merge_documents


def collect_docx_files(directory: Path = Path('.'), exclude_file: str = None) -> List[str]:
    """Collect all DOCX files in a directory, excluding output file."""
    files = []
    exclude_name = Path(exclude_file).name if exclude_file else None
    
    for f in directory.glob('*.docx'):
        if f.is_file() and not f.name.startswith('~$'):
            if exclude_name and f.name == exclude_name:
                continue
            files.append(str(f))
    
    return sorted(files, key=lambda x: Path(x).name.lower())


@click.command()
@click.option(
    '--inputs',
    multiple=True,
    help='Explicit list of .docx files to merge. If not provided, uses all .docx files in current directory.'
)
@click.option(
    '--output',
    default='all_documents.docx',
    help='Output file path/name. Default: all_documents.docx'
)
@click.option(
    '--menu-title',
    default='Menu',
    help='Heading text for the clickable menu. Default: Menu'
)
@click.option(
    '--back-label',
    default='Back to menu',
    help='Label for the backlink at the start of each section. Default: Back to menu'
)
@click.option(
    '--category-sep',
    default='_',
    help='Separator between category and document name in filenames. Default: _'
)
@click.option(
    '--dry-run',
    is_flag=True,
    default=False,
    help='Show what would be merged without writing output.'
)
def main(inputs, output, menu_title, back_label, category_sep, dry_run):
    """Merge multiple Word (.docx) files with a clickable navigation menu.
    
    If no input files are specified, automatically finds all .docx files 
    in the current directory (excluding the output file).
    """
    
    if inputs:
        input_files = list(inputs)
    else:
        input_files = collect_docx_files(exclude_file=output)
    
    if not input_files:
        click.echo("No input files found.", err=True)
        click.echo("Place .docx files in the current directory or specify files with --inputs", err=True)
        sys.exit(1)
    
    click.echo(f"Found {len(input_files)} files to merge:")
    for f in input_files:
        click.echo(f"  - {Path(f).name}")
    
    if dry_run:
        click.echo("\n--- DRY RUN MODE ---")
    
    try:
        result = merge_documents(
            input_files=input_files,
            output_path=output,
            menu_title=menu_title,
            back_label=back_label,
            category_separator=category_sep,
            toc_depth=2,
            keep_styles=True,
            dry_run=dry_run
        )
        
        if result:
            click.echo(f"\nSuccessfully created: {result}")
    except Exception as e:
        click.echo(f"Error: {e}", err=True)
        sys.exit(1)


if __name__ == '__main__':
    main()