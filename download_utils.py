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
    else:
        current_titles = {item["title"] for item in figs_meta}
        st.session_state[key] = {
            title: selected
            for title, selected in st.session_state[key].items()
            if title in current_titles
        }
        for item in figs_meta:
            st.session_state[key].setdefault(item["title"], True)
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
    Render chart selection checkboxes and one ZIP download button.
    """
    if not figs_meta:
        return

    selection_key = _ensure_selection_state(context_key, figs_meta)
    selection_dict = st.session_state[selection_key]

    st.caption("Wählen Sie die Diagramme aus, die in der ZIP-Datei enthalten sein sollen.")
    for item in figs_meta:
        title = item["title"]
        checkbox_key = f"dl_{context_key}_{item.get('filename', title)}"
        selection_dict[title] = st.checkbox(
            title,
            value=selection_dict.get(title, True),
            key=checkbox_key,
        )

    selected_count = sum(1 for selected in selection_dict.values() if selected)
    if selected_count == 0:
        st.warning("Bitte wählen Sie mindestens ein Diagramm für den Download aus.")

    # Create ZIP buffer
    buf = _figs_to_zip(figs_meta, selection_dict)

    # Show download button
    st.download_button(
        label=f"⬇️ Download ausgewählte Diagramme als ZIP ({selected_count})",
        data=buf,
        file_name=zip_name or f"charts_{context_key}.zip",
        mime="application/zip",
        key=f"zip_download_{context_key}",
        disabled=selected_count == 0,
    )
