def terminal_progress_bar(pct: float, text: str = "") -> None:
    """
    write progress bar in terminal
    :param pct: percent print
    :param text: (if need) text print after progress bar
    :return:
    """
    pct_f = float(pct)
    pct_prec = int(pct / 5.0)
    print(f"{pct_f:>7.2f}% [{'=' * pct_prec:>20}]", text, end="\r")
    return
