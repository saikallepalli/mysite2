pipeline {
    agent any
    stages {
        stage('Build image') {
            steps {
                sh 'docker build -t mysite2:latest .'
            }
        }
        stage('Deploy container') {
            steps {
                sh 'docker rm -f mysite2 || true'
                sh 'docker run -d --name mysite2 -p 9090:80 mysite2:latest'
            }
        }
    }
}