TEST_OUT=$(python /var/scripts/strahlzeiten.py -help 2>&1)

if !([ $? -eq 0 ]); then
    echo FAIL # goes to /var/mail/root
    echo -e "Subject: Calendar Script fwklux5 \n$TEST_OUT\n.\n" | /usr/sbin/sendmail p.petring@hzdr.de 
fi
