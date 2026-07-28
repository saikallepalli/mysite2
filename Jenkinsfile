pipeline {
    agent any
    stages {
        stage('Build image') {
            steps {
                sh 'docker build -t mysite:latest .'
            }
        }
        stage('Deploy container') {
            steps {
                sh 'docker rm -f mysite || true'
                sh 'docker run -d --name mysite -p 8090:80 mysite:latest'
            }
        }
    }
}