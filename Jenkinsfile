node('jenkins-linux') {
	properties([[$class: 'BuildDiscarderProperty',
                strategy: [$class: 'LogRotator', numToKeepStr: '5']],
                pipelineTriggers([[$class: 'BitBucketTrigger'], pollSCM(''), snapshotDependencies()])
            ])
        stage('Build') {
            def TAG_NAME = binding.variables.get("TAG_NAME")
            if (TAG_NAME != null) {
                sh "echo tag $TAG_NAME"
            } else {
                sh "echo Non-tag build"
            }
            withMaven(
                maven:'maven',

                 sh "mvn clean package"

      }
}
