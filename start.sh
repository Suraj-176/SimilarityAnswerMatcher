#!/bin/bash
conda env create -f environment.yml
conda activate similarity_matcher
exec streamlit run similarity_app.py