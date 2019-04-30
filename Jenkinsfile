Jenkinsfile ('jenkins-linux')
node {
    stage('Build') {
        echo 'Building....'
          def TAG_NAME = binding.variables.get("TAG_NAME")
                    if (TAG_NAME != null) {
                        sh "echo tag $TAG_NAME"
                    } else {
                        sh "echo Non-tag build"
                    }
            withMaven(
                           jdk:'jdk1.8-oracle',
                           maven:'InstalledMaven',
                           globalMavenSettingsConfig: '56ecb4c7-2efd-496d-949d-9209eee1c6a6',
                           ) {
                               sh "mvn clean package "
                           }



    }
    stage('Test') {
        echo 'Building....'
    }
    stage('Deploy') {
        echo 'Deploying....'
    }
}