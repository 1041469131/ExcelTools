node('jenkins-linux') {
	properties([[$class: 'BuildDiscarderProperty',
                strategy: [$class: 'LogRotator', numToKeepStr: '5']],
                pipelineTriggers([[$class: 'BitBucketTrigger'], pollSCM(''), snapshotDependencies()])
            ])
        stage('Build') {
            def TAG_NAME = binding.variables.get("TAG_NAME")
            if (TAG_NAME != null) {
                bat "echo tag $TAG_NAME"
            } else {
                bat "echo Non-tag build"
            }
            withMaven(
                jdk:'jdk1.8-oracle',
                maven:'InstalledMaven',
                globalMavenSettingsConfig: '56ecb4c7-2efd-496d-949d-9209eee1c6a6',
                ) {
                    bat "mvn clean package"
                }
      }
}
