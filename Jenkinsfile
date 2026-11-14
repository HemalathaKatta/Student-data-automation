pipeline {
    agent any

    triggers {
        // Check GitHub for new commits every 2 minutes
        pollSCM('H/2 * * * *')
    }

    stages {
        stage('Checkout') {
            steps {
                echo 'Downloading the latest files from GitHub...'
                checkout scm
                sh 'ls -l data incoming'
            }
        }

        stage('Check Python') {
            steps {
                sh '''
                    python3 --version
                    python3 -c "import openpyxl; print('openpyxl', openpyxl.__version__)"
                '''
            }
        }

        stage('Merge and Check Duplicates') {
            steps {
                echo 'Checking new records against old records and merging...'
                sh 'python3 merge_students.py'
            }
        }

        stage('Push Output to GitHub') {
            steps {
                withCredentials([usernamePassword(credentialsId: 'github-token',
                                                  usernameVariable: 'GIT_USER',
                                                  passwordVariable: 'GIT_TOKEN')]) {
                    sh '''
                        git config user.name "Jenkins"
                        git config user.email "jenkins@localhost"
                        git add output/
                        if git diff --cached --quiet; then
                            echo "Output unchanged - nothing to upload"
                        else
                            git commit -m "Jenkins: update merged_students.xlsx (build $BUILD_NUMBER)"
                            REPO=$(git config remote.origin.url | sed 's#https://##')
                            git pull --rebase "https://$GIT_USER:$GIT_TOKEN@$REPO" main
                            git push "https://$GIT_USER:$GIT_TOKEN@$REPO" HEAD:main
                            echo "merged_students.xlsx uploaded to GitHub"
                        fi
                    '''
                }
            }
        }
    }

    post {
        always {
            // The output file can also be downloaded from the build page in Jenkins
            archiveArtifacts artifacts: 'output/merged_students.xlsx', allowEmptyArchive: true
        }
        success { echo 'PIPELINE PASSED' }
        failure { echo 'PIPELINE FAILED: check the error above' }
    }
}
