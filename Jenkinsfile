Jenkinsfile ('jenkins-linux')
pipeline {
    agent any
    stages {
        stage('Build') {
            steps {
                echo 'Building'

                withMaven(
                                jdk:'jdk1.8-oracle',
                                maven:'InstalledMaven',
                                globalMavenSettingsConfig: '7e18ea53bead6cb35e6d4827192283a7cb04a6ec',
                                ) {
                                    sh "mvn clean package -DskipTests -U"
                                }

            }
        }
        stage('Test') {
            steps {
                echo 'Testing'
            }
        }
        stage('Deploy') {
            steps {
                echo 'Deploying'
            }
        }
    }
}