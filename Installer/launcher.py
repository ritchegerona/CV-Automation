import os
import sys
import shutil
import streamlit.web.bootstrap

if __name__ == "__main__":
    # Determine the directory where the bundle resources are located
    if getattr(sys, 'frozen', False):
        # Running inside the PyInstaller temporary extraction folder
        bundle_dir = sys._MEIPASS
    else:
        # Running in a normal development environment
        bundle_dir = os.path.dirname(os.path.abspath(__file__))
        # If running from inside the "Installer" directory in dev mode,
        # point bundle_dir back to the parent directory to find app.py, templates, etc.
        if os.path.basename(bundle_dir) == "Installer":
            bundle_dir = os.path.dirname(bundle_dir)

    # 1. Copy default template if it doesn't exist in current working directory
    template_name = "MSR_CV_Template.docx"
    target_template = os.path.join(os.getcwd(), template_name)
    if not os.path.exists(target_template):
        src_template = os.path.join(bundle_dir, template_name)
        if os.path.exists(src_template):
            try:
                shutil.copy(src_template, target_template)
                print(f"Copied default {template_name} to current directory.")
            except Exception as e:
                print(f"Warning: Could not copy default template: {e}")

    # Also copy the PDF template version if available
    pdf_template_name = "MSR_CV_Template.pdf"
    target_pdf_template = os.path.join(os.getcwd(), pdf_template_name)
    if not os.path.exists(target_pdf_template):
        src_pdf_template = os.path.join(bundle_dir, pdf_template_name)
        if os.path.exists(src_pdf_template):
            try:
                shutil.copy(src_pdf_template, target_pdf_template)
            except Exception as e:
                print(f"Warning: Could not copy PDF template: {e}")

    # 2. Create Candidate CV Summary folder if missing
    summary_dir = os.path.join(os.getcwd(), "Candidate CV Summary")
    if not os.path.exists(summary_dir):
        try:
            os.makedirs(summary_dir)
            print("Created 'Candidate CV Summary' directory.")
        except Exception as e:
            print(f"Warning: Could not create summary directory: {e}")

    # 3. Create a template .env file if it doesn't exist
    env_path = os.path.join(os.getcwd(), ".env")
    if not os.path.exists(env_path):
        try:
            with open(env_path, "w") as f:
                f.write("# Enter your Gemini API Key below\nGEMINI_API_KEY=your_gemini_api_key_here\n")
            print("Created a template .env file. Please edit it with your Gemini API key.")
        except Exception as e:
            print(f"Warning: Could not create template .env: {e}")

    # Set the config directory environment variable so Streamlit loads config.toml from the bundle
    os.environ["STREAMLIT_CONFIG_DIR"] = os.path.join(bundle_dir, ".streamlit")

    # Path to app.py inside the bundle
    app_path = os.path.join(bundle_dir, "app.py")

    # Override command line arguments to invoke Streamlit run
    sys.argv = ["streamlit", "run", app_path, "--server.headless=false"]

    # Boot Streamlit
    streamlit.web.bootstrap.run(
        main_script_path=app_path,
        is_hello=False,
        args=[],
        flag_options={}
    )
