#!/usr/bin/env python3
# -*- coding: utf-8 -*-


import platform, os, logging, sys
from pyupdater.client import Client
from client_config import ClientConfig


class autoUpdater():
    def __init__(self,queue,app_name, app_version):

        # OS
        PLATFORM = platform.system()  # -> 'Windows' / 'Linux' / 'Darwin'

        # SET WORKING DIRECTORY
        # This is required for the PATHS to work properly when app is frozen
        if PLATFORM == 'Windows':
            cwd = os.path.dirname(os.path.abspath(sys.argv[0]))
        elif PLATFORM == 'Linux':
            cwd = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        else:
            logging.warning(('This app has not been yet '
                            'tested for platform_system={}').format(PLATFORM))
            cwd = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
        os.chdir(cwd)


        #A callback to print download progress.
        def downloadStatus(info):
            total = info.get(u'total')
            downloaded = info.get(u'downloaded')
            progress = str(round(int(downloaded)*100/int(total)))
            status = info.get(u'status')

            if status == "finished":
                logging.info("Download finished")
            elif status == "downloading":
                logging.debug("Download progress: %s"%progress+"%")
                queue.put(progress)
            else:
                logging.warning("Unexpected download status: %s".format(status))

        #Initialize client with ClientConfig & later call refresh to get latest update data.
        #This HTTP key is needed to access to my repository.

        logging.debug("Opening client...")
        #client = Client(ClientConfig(),http = "reuMxY8PxCnrHctZyh29")
        client = Client(ClientConfig())
        client.refresh()

        #A hook to track download information
        client.add_progress_hook(downloadStatus)


        #Update_check returns an AppUpdate object if there is an update available
        logging.debug("Checking for updates...")
        app_update = client.update_check(app_name,app_version)


        #If we get an update object we can proceed to download the update.
        if app_update is not None:
            logging.info("A newer version was found")
            logging.debug("Downloading update...")
            app_update.download()
        else:
            logging.info("No update avalable")
            queue.put("NO_UPDATE")


        #Ensure file downloaded successfully, extract update, overwrite and restart current application
        if app_update is not None and app_update.is_downloaded():
            logging.debug("Overwriting current application...")
            app_update.extract_overwrite()
            logging.info("Succefuly overwrited")
            queue.put("DONE")

if __name__ == "__main__":
    # Logging configuration
    logging.basicConfig(level=logging.DEBUG, format='%(process)d-%(levelname)s-%(message)s')

