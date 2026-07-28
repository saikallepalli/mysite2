FROM jenkins/jenkins:lts
USER root
RUN apt-get update && apt-get install -y docker.io && rm -rf /var/lib/apt/lists/*
# let the jenkins user reach the Docker Desktop socket (owned by root)
RUN usermod -aG root jenkins
USER jenkins