"""Application entry point for the PCGD HỘ GIA ĐÌNH desktop UI."""

from gui import main
from multiprocessing import freeze_support


if __name__ == "__main__":
    freeze_support()
    main()
