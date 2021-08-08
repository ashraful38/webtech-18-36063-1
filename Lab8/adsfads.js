const form=document.getElementById('form');
const username=document.getElementById('username');
const password=document.getElementById('password');
const confirmPassword=document.getElementById('confirmPassword');
const email=document.getElementById('email');
const phoneNumber=document.getElementById('phoneNumber');
const phoneNumber=document.getElementById('phoneNumber');


//Show error message
function showError(input,message)
{
    const fromGroup=input.parentElement;
    fromGroup.className='form-group error';
    const small =fromGroup.querySelector('small');
    small.innerText=message;

}


//event listener

form.addEventListener('submit',function(e){
    e.preventDefault();

    if(username.value ==='')
    {
        showError(username,'Username is required');
    }
    else
    {
        showSuccess(username);
    }
})