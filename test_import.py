try:
    import tkcalendar
    print("tkcalendar import success")
    print(f"Location: {tkcalendar.__file__}")
except ImportError as e:
    print(f"Import failed: {e}")
