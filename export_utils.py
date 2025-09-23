# export_utils.py
import io
import zipfile


def fig_to_png_bytes(fig, dpi: int = 300, bbox: str = "tight") -> io.BytesIO:
    """
    Convert a Matplotlib figure into a PNG stored in memory.
    
    Args:
        fig: Matplotlib figure object.
        dpi: Resolution of the output PNG.
        bbox: Bounding box option for savefig.

    Returns:
        BytesIO object containing PNG data.
    """
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches=bbox)
    buf.seek(0)
    return buf


def bundle_zip(figs_meta, selection_dict: dict) -> io.BytesIO:
    """
    Bundle selected figures into an in-memory ZIP archive.

    Args:
        figs_meta: list of dicts {title, filename, fig}
        selection_dict: {title: bool} mapping of which figures to include

    Returns:
        BytesIO object with a ZIP containing selected PNGs.
    """
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
        for item in figs_meta:
            title = item["title"]
            if not selection_dict.get(title, False):
                continue
            fig = item["fig"]
            filename = item.get("filename", f"{title}.png")

            # Render to PNG and add to ZIP
            png_buf = fig_to_png_bytes(fig)
            zipf.writestr(filename, png_buf.read())

    zip_buf.seek(0)
    return zip_buf
