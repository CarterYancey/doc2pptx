import shutil
import tempfile
from pathlib import Path

import gradio as gr
import doc2pptx


APP_TITLE = "PPTX Generator"
APP_DESCRIPTION = (
    "Upload a source document and, optionally, a PowerPoint template. "
    "The server will return a generated .pptx file."
)


def _copy_with_original_name(src_path: str, dst_dir: str, fallback_name: str) -> Path:
    src = Path(src_path)
    name = src.name if src.name else fallback_name
    dst = Path(dst_dir) / name
    shutil.copy2(src, dst)
    return dst

def generate_pptx(
    document_file,
    template_file,
    use_llm,
    ollama_host,
    ollama_model,
    custom_prompt,
    max_chunk_chars,
    chunk_overlap_sentences,
    verbose_logging,
):
    if document_file is None:
        raise gr.Error("Please upload a document file.")

    workdir = Path(tempfile.mkdtemp(prefix="claude_doc2pptx_"))
    log_path = workdir / "doc2pptx.log"

    # Always write a fresh log file for each run; verbose toggles terminal output only.
    doc2pptx.configure_logging(
        log_file=log_path,
        verbose=bool(verbose_logging),
        quiet=False,
    )

    try:
        input_doc = _copy_with_original_name(
            document_file,
            str(workdir),
            "uploaded_document"
        )
        output_pptx = workdir / f"{input_doc.stem}.pptx"

        input_template = None
        if template_file is not None:
            input_template = _copy_with_original_name(
                template_file,
                str(workdir),
                "uploaded_template.pptx"
            )

        prompt_file_path = None
        if custom_prompt and custom_prompt.strip() and custom_prompt.strip() != doc2pptx.DEFAULT_REWRITE_PROMPT.strip():
            prompt_path = workdir / "custom_prompt.txt"
            prompt_path.write_text(custom_prompt, encoding="utf-8")
            prompt_file_path = str(prompt_path)

        doc2pptx.generate_pptx(
            input_path=str(input_doc),
            output_path=str(output_pptx),
            template_path=str(input_template) if input_template else None,
            title=None,
            use_llm=bool(use_llm),
            ollama_host=(ollama_host or doc2pptx.DEFAULT_OLLAMA_HOST).strip(),
            ollama_model=(ollama_model or doc2pptx.DEFAULT_OLLAMA_MODEL).strip(),
            llm_prompt_file=prompt_file_path,
            max_chunk_chars=int(max_chunk_chars) if max_chunk_chars else 6000,
            chunk_overlap_sentences=int(chunk_overlap_sentences) if chunk_overlap_sentences is not None else 2,
        )

        if not output_pptx.exists():
            raise gr.Error("The generator finished, but output.pptx was not created.")

        return str(output_pptx), str(log_path)
    except Exception as exc:
        # Surface logs even on failure so the user can diagnose.
        raise gr.Error(str(exc)) from exc


with gr.Blocks(title=APP_TITLE) as demo:
    gr.Markdown(f"# {APP_TITLE}")
    gr.Markdown(APP_DESCRIPTION)

    with gr.Row():
        document_input = gr.File(
            label="Source document",
            file_count="single",
            type="filepath",
        )
        template_input = gr.File(
            label="Optional PowerPoint template (.pptx)",
            file_count="single",
            type="filepath",
        )

    with gr.Accordion("LLM rewrite (local Ollama)", open=False):
        use_llm_input = gr.Checkbox(
            label="Rewrite extracted text with a local LLM before building slides",
            value=True,
        )
        with gr.Row():
            ollama_host_input = gr.Textbox(
                label="Ollama host",
                value=doc2pptx.DEFAULT_OLLAMA_HOST,
            )
            ollama_model_input = gr.Textbox(
                label="Ollama model",
                value=doc2pptx.DEFAULT_OLLAMA_MODEL,
            )
        custom_prompt_input = gr.Textbox(
            label="System prompt (edit to override; restore default by matching the text below)",
            value=doc2pptx.DEFAULT_REWRITE_PROMPT,
            lines=10,
        )
        with gr.Row():
            max_chunk_chars_input = gr.Number(
                label="Max chunk chars",
                value=20000,
                precision=0,
                info="Large documents are always split on headings/paragraphs. Lower this for small local models.",
            )
            chunk_overlap_sentences_input = gr.Number(
                label="Chunk overlap sentences",
                value=2,
                precision=0,
                info="Trailing sentences from the previous chunk prepended to the next for continuity.",
            )


    run_button = gr.Button("Generate PowerPoint", variant="primary")

    output_file = gr.File(label="Generated PowerPoint")
    with gr.Accordion("Logging", open=False):
        verbose_logging_input = gr.Checkbox(
            label="Verbose terminal output (DEBUG) — includes full extracted text, chunk contents, and LLM prompts/responses in the server log",
            value=False,
        )
        gr.Markdown(
            "A detailed log (`doc2pptx.log`) is always written to the run workdir and shown below after each run. "
            "It contains the full extracted text, every chunk, every LLM prompt/input/response, and per-chunk timings."
        )
    log_file_output = gr.File(label="Log File (download the full trace)")

    run_button.click(
        fn=generate_pptx,
        inputs=[
            document_input,
            template_input,
            use_llm_input,
            ollama_host_input,
            ollama_model_input,
            custom_prompt_input,
            max_chunk_chars_input,
            chunk_overlap_sentences_input,
            verbose_logging_input,
        ],
        outputs=[output_file, log_file_output],
    )


if __name__ == "__main__":
    demo.launch(server_name="0.0.0.0", server_port=7860)

