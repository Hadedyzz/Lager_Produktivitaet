import io
import zipfile
import streamlit as st


def _ensure_selection_state(context_key: str, figs_meta):
    """
    Ensure st.session_state has a persistent selection dict
    for this context_key. Initialize to all True if new.
    """
    key = f"download_selection_{context_key}"
    if key not in st.session_state:
        st.session_state[key] = {item["title"]: True for item in figs_meta}
    return key


def _figs_to_zip(figs_meta, selection_dict):
    """
    Create an in-memory ZIP archive of selected figures.
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zipf:
        for item in figs_meta:
            title = item["title"]
            if not selection_dict.get(title, False):
                continue
            fig = item["fig"]
            filename = item.get("filename", f"{title}.png")

            # Render fig to PNG
            png_buf = io.BytesIO()
            fig.savefig(png_buf, format="png", dpi=300, bbox_inches="tight")
            png_buf.seek(0)

            zipf.writestr(filename, png_buf.read())
    buf.seek(0)
    return buf


def render_download_section(figs_meta, context_key: str, zip_name: str = None):
    """
    Render a simple download section with only one button (no checkboxes).
    """
    if not figs_meta:
        return

    # Always include all figures
    selection_dict = {item["title"]: True for item in figs_meta}

    # Create ZIP buffer
    buf = _figs_to_zip(figs_meta, selection_dict)

    # Show download button
    st.download_button(
        label=f"⬇️ Download alle Diagramme als ZIP",
        data=buf,
        file_name=zip_name or f"charts_{context_key}.zip",
        mime="application/zip",
        key=f"zip_download_{context_key}",
    )
