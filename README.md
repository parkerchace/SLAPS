# SLAPS
Solution for LAPS: a new gui tool for enterprises or individuals with Windows AD domains
To add:

# Add expiration date display  
$searcher.PropertiesToLoad.Add("ms-Mcs-AdmPwdExpirationTime") | Out-Null  
$expiryDate = [datetime]::FromFileTime($result.Properties["ms-Mcs-AdmPwdExpirationTime"][0])  
