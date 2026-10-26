"""
Streamlit frontend for the PO Translation System.

Lets a user upload one or more workbooks, translate them to Tamil and Hindi
through the FastAPI backend, and download the results.
"""

import io
import os
import zipfile
from pathlib import Path

import requests
import streamlit as st

API_BASE_URL = os.getenv("API_BASE_URL", "http://localhost:8000").rstrip("/")
UPLOAD_TIMEOUT = 60
TRANSLATE_TIMEOUT = 600
XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def get_error_message(response: requests.Response) -> str:
    """Read the error detail returned by the API."""
    try:
        return str(response.json().get("detail", response.text))
    except ValueError:
        return response.text or f"HTTP {response.status_code}"


def translate_po(uploaded_file) -> dict:
    """Run upload, translate and download against the API. Returns the result."""
    with st.status(f"Translating {uploaded_file.name}...", expanded=True) as progress:
        st.write("Uploading file")
        response = requests.post(
            f"{API_BASE_URL}/api/upload",
            files={"file": (uploaded_file.name, uploaded_file.getvalue(), XLSX_MIME_TYPE)},
            timeout=UPLOAD_TIMEOUT,
        )
        if response.status_code != 200:
            progress.update(label="Upload failed", state="error")
            raise RuntimeError(get_error_message(response))
        file_id = response.json()["file_id"]

        st.write("Translating to Tamil and Hindi")
        response = requests.post(f"{API_BASE_URL}/api/translate/{file_id}", timeout=TRANSLATE_TIMEOUT)
        if response.status_code != 200:
            progress.update(label="Translation failed", state="error")
            raise RuntimeError(get_error_message(response))
        result = response.json()

        st.write("Preparing download")
        response = requests.get(f"{API_BASE_URL}/api/download/{file_id}", timeout=UPLOAD_TIMEOUT)
        if response.status_code != 200:
            progress.update(label="Download failed", state="error")
            raise RuntimeError(get_error_message(response))

        progress.update(label=f"{uploaded_file.name} translated", state="complete", expanded=False)

    return {
        "source_name": uploaded_file.name,
        "output_filename": result["output_filename"],
        "languages": result["languages"],
        "translated_cells": result.get("translated_cells", 0),
        "content": response.content,
    }


def build_zip(results: list) -> bytes:
    """Pack all translated workbooks into one zip file."""
    buffer = io.BytesIO()
    used_names = set()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for result in results:
            # Two uploads with the same name must not overwrite each other
            name = result["output_filename"]
            counter = 2
            while name in used_names:
                name = f"{Path(result['output_filename']).stem}_{counter}.xlsx"
                counter += 1
            used_names.add(name)
            archive.writestr(name, result["content"])
    return buffer.getvalue()


st.set_page_config(page_title="SOP Translation System", page_icon="📄", layout="centered")

st.title("SOP Translation System")


uploaded_files = st.file_uploader("SOP workbooks", type=["xlsx"], accept_multiple_files=True)

# The selection changed: forget the previous results
selection = [(file.name, file.size) for file in uploaded_files]
if st.session_state.get("selection") != selection:
    st.session_state["selection"] = selection
    st.session_state.pop("results", None)
    st.session_state.pop("errors", None)

if uploaded_files and st.button("Translate", type="primary"):
    results, errors = [], []
    for uploaded_file in uploaded_files:
        # One file that fails must not stop the others
        try:
            results.append(translate_po(uploaded_file))
        except requests.exceptions.ConnectionError:
            errors.append(
                f"Cannot reach the API at {API_BASE_URL}. "
                "Start the application with: python run.py"
            )
            break
        except requests.exceptions.Timeout:
            errors.append(f"{uploaded_file.name}: The request timed out. Please try again.")
        except RuntimeError as e:
            errors.append(f"{uploaded_file.name}: {e}")
    st.session_state["results"] = results
    st.session_state["errors"] = errors

for error in st.session_state.get("errors", []):
    st.error(error)

results = st.session_state.get("results", [])
if results:
    st.success(f"Translated {len(results)} of {len(uploaded_files)} files.")
    if len(results) > 1:
        st.download_button(
            label="Download all (.zip)",
            data=build_zip(results),
            file_name="Translated_SOPs.zip",
            mime="application/zip",
            type="primary",
        )
    for index, result in enumerate(results):
        st.download_button(
            label=f"Download {result['output_filename']}",
            data=result["content"],
            file_name=result["output_filename"],
            mime=XLSX_MIME_TYPE,
            key=f"download_{index}",
            help=f"{result['translated_cells']} cells translated to "
                 f"{' and '.join(result['languages'])}",
        )
