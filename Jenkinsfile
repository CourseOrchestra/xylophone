@Library('ratcheting') _
node {
    def server = Artifactory.server 'ART'
    def rtMaven = Artifactory.newMavenBuild()
    def buildInfo
    def oldWarnings

    stage ('Clone') {
        checkout scm
    }

    stage ('Artifactory configuration') {
        rtMaven.tool = 'M3' 
        rtMaven.deployer releaseRepo: 'libs-release-local', snapshotRepo: 'libs-snapshot-local', server: server
        rtMaven.resolver releaseRepo: 'libs-release', snapshotRepo: 'libs-snapshot', server: server
        buildInfo = Artifactory.newBuildInfo()
        buildInfo.env.capture = true


        def downloadSpec = """
                 {"files": [
                    {
                      "pattern": "warn/xylophone/*/warnings.yml",
                      "build": "xylophone :: dev/LATEST",
                      "target": "previous.yml",
                      "flat": "true"
                    }
                    ]
                }"""
        server.download spec: downloadSpec
        oldWarnings = readYaml file: 'previous.yml'
    }


    stage ('Spellcheck'){
        result = sh (returnStdout: true,
           script: """for f in \$(find . -name '*.adoc'); do cat \$f | sed "s/-/ /g" | aspell --master=en --personal=./dict list; done | sort | uniq""")
              .trim()
        if (result) {
           echo "The following words are probaly misspelled:"
           echo result
           error "Please correct the spelling or add the words above to the local dictionary."
        }
    }

    try{
        stage ('Exec Maven') {
            rtMaven.run pom: 'pom.xml', goals: 'clean install', buildInfo: buildInfo
        }
    } finally {
        junit 'target/surefire-reports/**/*.xml'
        recordIssues tool: checkStyle(pattern: 'target/checkstyle-result.xml')
        recordIssues tool: spotBugs(pattern: 'target/spotbugsXml.xml')
    }

    stage ('Ratcheting') {
            def warningsMap = countWarnings()
            writeYaml file: 'target/warnings.yml', data: warningsMap
            compareWarningMaps oldWarnings, warningsMap
    }

    if (env.BRANCH_NAME == 'dev') {
        stage ('Publish build info') {

            def uploadSpec = """
                        {
                         "files": [
                            {
                              "pattern": "target/warnings.yml",
                              "target": "warn/xylophone/${currentBuild.number}/warnings.yml"
                            }
                            ]
                        }"""
            def buildInfo2 = server.upload spec: uploadSpec
            buildInfo.append(buildInfo2)
            server.publishBuildInfo buildInfo
        }
    }
}
