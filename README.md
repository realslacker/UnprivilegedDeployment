# UnprivilegedDeployment
Run an installation as an unprivileged user.

The goal of this project is to have a installation launch as a privileged user and interact with an unprivileged user.

# Example Setup

1. Create a scheduled task that runs as the SYSTEM account at startup. That scheduled task should import the module, and then run the **Start-UnprivilegedDeploymentInstaller**. The installer process will wait for the Client process to run and start the installation.
2. Create a scheduled task that runs as the logged on user at logon. That scheduled task should import the module, and then run the **Start-UnprivilegedDeploymentClient**. When the user is ready they can kick off the installation.

# Submitting Patches / Issues

Please feel free to submit patches, errors, or language packs.
