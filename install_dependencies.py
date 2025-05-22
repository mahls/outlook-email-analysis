import subprocess
import sys
import platform

def install(package):
    """
    Installs a given Python package using pip.
    """
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"Successfully installed {package}")
    except subprocess.CalledProcessError as e:
        print(f"Failed to install {package}: {e}")
        raise # Re-raise the exception to stop if a core package fails

def main():
    """
    Main function to install all required Python packages and download necessary NLTK/Stanza data.
    """
    print("Starting installation of required Python packages...")

    # Install numpy first as it's a common dependency for many scientific packages
    try:
        install("numpy")
    except Exception as e:
        print(f"Initial attempt to install numpy failed. This might cause issues for other packages. Error: {e}")
        # Continue, as some systems might have it pre-installed or other issues might resolve later.

    # List of core dependencies for the email analysis dashboard
    packages = [
        "streamlit",
        "pandas",
        "matplotlib",
        "seaborn",
        "wordcloud",
        "textblob",
        "stanza",
        "plotly",
        "scikit-learn",
        "tqdm", # Often useful for progress bars, though not directly used in the dashboard logic itself
        "nltk"
    ]

    for package in packages:
        try:
            install(package)
        except Exception as e:
            print(f"Could not install {package}. Please check the error message above for details. "
                  "You might need to install build tools (e.g., Visual C++ Build Tools on Windows) "
                  "or resolve network issues.")

    # Download TextBlob corpora
    print("\nAttempting to download TextBlob corpora...")
    try:
        import textblob.download_corpora as tb_download
        tb_download.download_all()
        print("✅ TextBlob corpora downloaded successfully.")
    except Exception as e:
        print(f"Error downloading TextBlob corpora. This might affect sentiment analysis. Error: {e}")

    # Download stanza English model
    print("\nAttempting to download Stanza English model...")
    try:
        import stanza
        # Check if model is already downloaded to avoid re-downloading
        # This check is a heuristic; actual download might still occur if corrupted/incomplete
        if not os.path.exists(os.path.join(stanza.DEFAULT_MODEL_DIR, 'en', 'default.pt')):
            stanza.download('en')
        else:
            print("Stanza English model already exists. Skipping download.")
        print("✅ Stanza English model ready.")
    except Exception as e:
        print(f"Error downloading Stanza English model. This will affect Named Entity Recognition. Error: {e}")

    print("\n✅ All core dependencies installation attempts complete.")
    print("You can now run your Streamlit email analysis dashboard with:")
    if platform.system() == "Windows":
        print("   streamlit run outlook-email-analysis-dashboard.py") # Generic filename
    else:
        print("   streamlit run ./outlook-email-analysis-dashboard.py") # Generic filename

if __name__ == "__main__":
    main()
