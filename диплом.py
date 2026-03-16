from dotenv import load_dotenv

from rentability import RentabilityAnalysisApp


if __name__ == "__main__":
    load_dotenv()
    app = RentabilityAnalysisApp()
    if app.winfo_exists():
        app.mainloop()
