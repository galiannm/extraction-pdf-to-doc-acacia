from pathlib import Path

BASE_DIR = Path(__file__).parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
EXTRACTED_DIR = OUTPUT_DIR / "extracted"
ASSETS_DIR = BASE_DIR / "assets"
GLOSSARY_PATH = BASE_DIR / "glossary.json"

LOGO_PATH = ASSETS_DIR / "logo-acacia.jpeg"
LOGO_SMALL_PATH = ASSETS_DIR / "logo-A.jpeg"

# Source PDFs — order matters for glossary building
HMH_DIR = BASE_DIR.parent / "HMH"
SOURCE_PDFS = [
    HMH_DIR / "MHM PS MS - Teachers' guide (K1-K2) intro.pdf",
    HMH_DIR / "MHM PS MS - Teachers' guide (K1-K2) period 1.pdf",
    HMH_DIR / "MHM PS MS - Teachers' guide (K1-K2) period 2.pdf",
    HMH_DIR / "MHM PS MS - Teachers' guide (K1-K2) period 3.pdf",
    HMH_DIR / "MHM PS MS - Teachers' guide (K1-K2) period 4.pdf",
    HMH_DIR / "MHM PS MS - Teachers' guide (K1-K2) period 5.pdf",
]
TEST_PDF = SOURCE_PDFS[0]

# Acacia brand colors as (R, G, B) tuples
class Colors:
    BLUE   = (91,  200, 245)   # A — sky blue
    ORANGE = (245, 166,  35)   # C — warm orange
    YELLOW = (248, 200,  64)   # A — golden yellow
    CORAL  = (232, 120, 106)   # C — coral / salmon
    GREEN  = (141, 198,  63)   # sprout — fresh green
    PURPLE = (176, 124, 198)   # A + bubble — soft purple

    WHITE      = (255, 255, 255)
    DARK_GREY  = (64,   64,  64)
    CREAM      = (250, 245, 238)
    LIGHT_GREY = (240, 240, 240)

    # Document structure mapping
    PERIOD_COVER            = ORANGE
    RITUALISED_ACTIVITIES   = YELLOW   # Activités ritualisées column
    GUIDED_LEARNING         = BLUE     # Apprentissages guidés column
    AUTONOMOUS_ACTIVITIES   = GREEN    # Autonomie et semi-autonomie column
    OBJECTIVES_HEADER       = PURPLE
    SECTION_BANNER          = CORAL
    PS_BADGE                = BLUE
    MS_BADGE                = ORANGE
    WEEK_ROW                = ORANGE   # SEMAINE N row background

class Fonts:
    HEADING  = "Mali"
    BODY     = "Nunito"
