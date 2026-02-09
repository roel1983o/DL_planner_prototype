from __future__ import annotations

import os
import uuid
from dataclasses import dataclass
from typing import Optional

from .parser_core import parse_excel_story_list, save_verhalenaanbod_xlsx
from .deel1_core import run_deel1
from .deel2_core import generate_handout_pdf

@dataclass
class PipelineOutputs:
    job_id: str
    deel1_xlsx: str
    deel2_pdf: str

def run_pipeline(*, uploaded_parser_input: str, workdir: str,
                 assets_dir: str) -> PipelineOutputs:
    os.makedirs(workdir, exist_ok=True)
    job_id = os.path.basename(workdir)

    # 1) PARSER
    df = parse_excel_story_list(uploaded_parser_input)

    # save_verhalenaanbod_xlsx schrijft standaard naar current working directory
    cwd = os.getcwd()
    os.chdir(workdir)
    try:
        fname = save_verhalenaanbod_xlsx(df)  # retourneert bestandsnaam, bv. "verhalenaanbod.xlsx"
    finally:
        os.chdir(cwd)

    verhalenaanbod_path = os.path.join(workdir, fname)

    # 2) DEEL 1
    excel_assets = os.path.join(assets_dir, "excel")
    deel1_out = os.path.join(workdir, "Krantenplanning.xlsx")
    run_deel1(
        templates_path=os.path.join(excel_assets, "Templates.xlsx"),
        beslispad_spread_path=os.path.join(excel_assets, "Beslispad Spread.xlsx"),
        beslispad_ep_path=os.path.join(excel_assets, "Beslispad EP.xlsx"),
        posities_path=os.path.join(excel_assets, "Posities en kenmerken.xlsx"),
        verhalenaanbod_path=verhalenaanbod_path,
        out_path=deel1_out,
    )

    # 3) DEEL 2
    template_dir = os.path.join(assets_dir, "images", "templates", "public")
    logo_file = os.path.join(assets_dir, "images", "Logo.jpg")
    deel2_out = os.path.join(workdir, "Krantenplanning_handout.pdf")
    generate_handout_pdf(deel1_out, template_dir, logo_file, out_pdf=deel2_out)

    return PipelineOutputs(job_id=job_id, deel1_xlsx=deel1_out, deel2_pdf=deel2_out)
