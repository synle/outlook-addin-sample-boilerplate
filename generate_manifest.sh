# {addin_app_uuid}
# {addin_app_url}
# 7164e750-dc86-49c0-b548-1bac57abdc7c
# https://localhost.com
file_manifest=manifest.xml
cp manifest.sample.xml manifest.xml
echo "generate manifest: $file_manifest"
perl -pi -e 's/{addin_app_uuid}/7164e750-dc86-49c0-b548-1bac57abdddd/g' $file_manifest
perl -pi -e 's/{addin_app_url}/https:\/\/localhost.com/g' $file_manifest
perl -pi -e 's/{addin_app_title}/My Outlook Demo Addin/g' $file_manifest
perl -pi -e 's/{addin_app_description}/Demo Addin Description/g' $file_manifest
