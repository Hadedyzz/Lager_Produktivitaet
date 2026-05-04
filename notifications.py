import streamlit as st


MESSAGES_KEY = "app_messages"


def reset_messages():
    """Clear collected user-facing messages for a new workbook context."""
    st.session_state[MESSAGES_KEY] = []


def add_message(level: str, text: str):
    """Collect a non-blocking message for the shared Meldungen expander."""
    if MESSAGES_KEY not in st.session_state:
        st.session_state[MESSAGES_KEY] = []

    message = {"level": level, "text": str(text)}
    if message not in st.session_state[MESSAGES_KEY]:
        st.session_state[MESSAGES_KEY].append(message)


def render_messages():
    """Render collected messages as compact lines inside one expander."""
    messages = st.session_state.get(MESSAGES_KEY, [])
    if not messages:
        return

    labels = {
        "info": "Info",
        "warning": "Warnung",
        "error": "Fehler",
    }
    with st.expander(f"Meldungen ({len(messages)})", expanded=False):
        for message in messages:
            label = labels.get(message.get("level"), "Meldung")
            st.markdown(f"- **{label}:** {message.get('text', '')}")
