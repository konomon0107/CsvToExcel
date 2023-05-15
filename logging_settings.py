import logging

def logging_start():
    
    logger = logging.getLogger('logging_settings')
    logger.setLevel(logging.DEBUG)
    # Create A File Handler With "example.log" As The Output Destination
    ch = logging.FileHandler(filename="logging.log")
    # View Up To DEBUG Level
    ch.setLevel(logging.DEBUG)
    # Log Description Format
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    # Give Formatting Information To File Handlers
    ch.setFormatter(formatter)
    # Give The Logger (Instance) Information About The File Handler
    logger.addHandler(ch)