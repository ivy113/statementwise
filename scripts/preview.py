from rich import print
from rich.console import Console, Group
from rich.panel import Panel
from rich.table import Table
from loguru import logger

from statementwise.cc_parser import AmexStatementParser

if __name__ == "__main__":
    logger.remove()
    amex_file = AmexStatementParser("data/1001.xlsx",
                                    "Transaction Details")
    amex_summary, amex_transactions = amex_file.parse()
    # print("\n --- Final Amex Summary ---")
    # print(amex_summary)
    
    # print("\n --- Final Amex Transactions ---")
    # print(amex_transactions.head())
    
    # --- Format and print the summary beautifully ---
    console = Console()

    # Create a table to hold the summary details
    summary_table = Table(show_header=False, box=None, padding=(0, 2))
    summary_table.add_column("Key", style="cyan")
    summary_table.add_column("Value", style="magenta")

    # Populate the table from your dictionary
    for key, value in amex_summary.items():
        # Make keys more readable (e.g., 'account_number' -> 'Account Number')
        display_key = key.replace('_', ' ').title()
        summary_table.add_row(f"[bold]{display_key}[/bold]", str(value))

    transactions_preview = amex_transactions.head()
    transactions_table = Table(title="[bold]Transactions Preview[/bold]", show_lines=True, expand=True)
    
    for column in transactions_preview.columns:
        transactions_table.add_column(column)
    
    for row in transactions_preview.iter_rows():
        transactions_table.add_row(*[str(item) for item in row])

    content_group = Group(
        summary_table,
        transactions_table
    )
    panel = Panel(
        content_group,
        title="[bold bright_white]💳 AMEX Statement[/bold bright_white]",
        border_style="green",
        padding=(1,2)
    )
    console.print(panel)
    # Wrap the table in a styled panel
